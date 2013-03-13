VERSION 5.00
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form FrmPlanillaCobranza 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Planilla de Cobranza"
   ClientHeight    =   8190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13530
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8190
   ScaleWidth      =   13530
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   7845
      Left            =   150
      TabIndex        =   2
      Top             =   120
      Width           =   13215
      _ExtentX        =   23310
      _ExtentY        =   13838
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   16777215
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Datos Generales"
      TabPicture(0)   =   "FrmPlanillaCobranza.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Detalle"
      TabPicture(1)   =   "FrmPlanillaCobranza.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame7"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "frmbotones"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0C0C0&
         Height          =   3015
         Left            =   -72810
         TabIndex        =   52
         Top             =   1770
         Width           =   7695
         Begin VB.Frame Frame1 
            BackColor       =   &H00FF8080&
            BorderStyle     =   0  'None
            Height          =   2655
            Left            =   150
            TabIndex        =   53
            Top             =   240
            Width           =   7425
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
               Left            =   3450
               TabIndex        =   54
               Top             =   780
               Width           =   1185
               _ExtentX        =   2090
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   10
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuda1 
               Height          =   285
               Left            =   3450
               TabIndex        =   55
               Top             =   270
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
            Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuda2 
               Height          =   285
               Left            =   3480
               TabIndex        =   56
               Top             =   1350
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
            Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuempresa 
               Height          =   315
               Left            =   3480
               TabIndex        =   60
               Top             =   1920
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
               Left            =   960
               TabIndex        =   61
               Top             =   1920
               Width           =   2085
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "TIPO DE PLANILLA"
               BeginProperty Font 
                  Name            =   "System"
                  Size            =   9.75
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
               TabIndex        =   59
               Top             =   300
               Width           =   2085
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "FECHA DE COBRANZA"
               BeginProperty Font 
                  Name            =   "System"
                  Size            =   9.75
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
               TabIndex        =   58
               Top             =   840
               Width           =   2385
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "OFICINA / VENDEDOR"
               BeginProperty Font 
                  Name            =   "System"
                  Size            =   9.75
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
               TabIndex        =   57
               Top             =   1380
               Width           =   2205
            End
         End
      End
      Begin VB.Frame Frame4 
         Height          =   930
         Left            =   -69300
         TabIndex        =   39
         Top             =   5430
         Width           =   1980
         Begin VB.CommandButton cmdBotones 
            Caption         =   "&Cancelar"
            Height          =   690
            Index           =   12
            Left            =   1050
            Picture         =   "FrmPlanillaCobranza.frx":0038
            Style           =   1  'Graphical
            TabIndex        =   1
            Top             =   180
            Width           =   855
         End
         Begin VB.CommandButton cmdBotones 
            Caption         =   "&Acepta"
            Height          =   690
            Index           =   11
            Left            =   90
            Picture         =   "FrmPlanillaCobranza.frx":047A
            Style           =   1  'Graphical
            TabIndex        =   0
            Top             =   180
            Width           =   870
         End
      End
      Begin VB.Frame frmbotones 
         Height          =   930
         Left            =   4890
         TabIndex        =   34
         Top             =   6690
         Width           =   4290
         Begin VB.CommandButton cmdBotones 
            Caption         =   "&Salir"
            Height          =   690
            Index           =   4
            Left            =   3255
            Picture         =   "FrmPlanillaCobranza.frx":08BC
            Style           =   1  'Graphical
            TabIndex        =   38
            Top             =   180
            Width           =   870
         End
         Begin VB.CommandButton cmdBotones 
            Caption         =   "&Eliminar"
            Height          =   690
            Index           =   2
            Left            =   2280
            Picture         =   "FrmPlanillaCobranza.frx":0CFE
            Style           =   1  'Graphical
            TabIndex        =   37
            Top             =   180
            Width           =   825
         End
         Begin VB.CommandButton cmdBotones 
            Caption         =   "&Grabar"
            Height          =   690
            Index           =   1
            Left            =   1260
            Picture         =   "FrmPlanillaCobranza.frx":1140
            Style           =   1  'Graphical
            TabIndex        =   36
            Top             =   180
            Width           =   870
         End
         Begin VB.CommandButton cmdBotones 
            Caption         =   "&Nuevo"
            Height          =   690
            Index           =   0
            Left            =   180
            Picture         =   "FrmPlanillaCobranza.frx":1582
            Style           =   1  'Graphical
            TabIndex        =   35
            Top             =   180
            Width           =   870
         End
      End
      Begin VB.Frame Frame7 
         Height          =   885
         Left            =   270
         TabIndex        =   28
         Top             =   6660
         Width           =   1335
         Begin VB.CommandButton cContado 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Documento Contado"
            BeginProperty Font 
               Name            =   "System"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   705
            Left            =   60
            Style           =   1  'Graphical
            TabIndex        =   29
            Top             =   150
            Width           =   1245
         End
      End
      Begin VB.Frame Frame3 
         Height          =   6045
         Left            =   210
         TabIndex        =   3
         Top             =   540
         Width           =   12855
         Begin VB.Frame Frame8 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   4845
            Left            =   0
            TabIndex        =   30
            Top             =   60
            Visible         =   0   'False
            Width           =   12765
            Begin TrueOleDBGrid70.TDBGrid DGrid1 
               Height          =   3885
               Left            =   360
               TabIndex        =   33
               Top             =   240
               Width           =   11985
               _ExtentX        =   21140
               _ExtentY        =   6853
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
               _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=117,.bold=0,.fontsize=825,.italic=0"
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
            Begin VB.CommandButton cCerrar 
               BackColor       =   &H0000C0C0&
               Caption         =   "&Cerrar"
               Height          =   465
               Left            =   10890
               Style           =   1  'Graphical
               TabIndex        =   32
               Top             =   4185
               Width           =   1170
            End
            Begin VB.CommandButton cAcepto 
               BackColor       =   &H0000C0C0&
               Caption         =   "&Acepta"
               Height          =   465
               Left            =   9660
               Style           =   1  'Graphical
               TabIndex        =   31
               Top             =   4185
               Width           =   1170
            End
            Begin VB.Shape Shape1 
               BorderColor     =   &H0000C0C0&
               BorderWidth     =   3
               FillColor       =   &H00FFFFFF&
               FillStyle       =   0  'Solid
               Height          =   4785
               Left            =   60
               Shape           =   4  'Rounded Rectangle
               Top             =   30
               Width           =   12585
            End
         End
         Begin VB.Frame Frame5 
            Height          =   1335
            Left            =   150
            TabIndex        =   5
            Top             =   4650
            Width           =   12555
            Begin VB.CommandButton cAyuda 
               Caption         =   "..."
               Height          =   315
               Index           =   4
               Left            =   7200
               TabIndex        =   62
               Top             =   480
               Width           =   195
            End
            Begin VB.CommandButton cAyuda 
               Caption         =   "..."
               Height          =   315
               Index           =   3
               Left            =   9300
               TabIndex        =   16
               Top             =   450
               Width           =   195
            End
            Begin VB.CommandButton cAyuda 
               Caption         =   "..."
               Height          =   315
               Index           =   2
               Left            =   1350
               TabIndex        =   27
               Top             =   450
               Width           =   150
            End
            Begin VB.CommandButton cAyuda 
               Caption         =   "..."
               Height          =   315
               Index           =   1
               Left            =   4980
               TabIndex        =   26
               Top             =   450
               Width           =   195
            End
            Begin VB.CommandButton cAyuda 
               Caption         =   "..."
               Height          =   315
               Index           =   0
               Left            =   3870
               TabIndex        =   25
               Top             =   450
               Width           =   195
            End
            Begin VB.TextBox Text1 
               Height          =   285
               Index           =   11
               Left            =   10200
               MaxLength       =   8
               TabIndex        =   18
               Top             =   450
               Width           =   1065
            End
            Begin VB.TextBox Text1 
               Height          =   285
               Index           =   10
               Left            =   11340
               MaxLength       =   12
               TabIndex        =   20
               Top             =   450
               Width           =   1065
            End
            Begin VB.TextBox Text1 
               Height          =   285
               Index           =   9
               Left            =   8820
               MaxLength       =   2
               TabIndex        =   15
               Top             =   450
               Width           =   435
            End
            Begin VB.TextBox Text1 
               Height          =   285
               Index           =   8
               Left            =   9630
               MaxLength       =   2
               TabIndex        =   17
               Top             =   450
               Width           =   435
            End
            Begin VB.TextBox Text1 
               Height          =   285
               Index           =   7
               Left            =   5910
               MaxLength       =   10
               TabIndex        =   13
               Top             =   450
               Width           =   1245
            End
            Begin VB.TextBox Text1 
               Height          =   285
               Index           =   0
               Left            =   90
               MaxLength       =   11
               TabIndex        =   6
               Top             =   450
               Width           =   1275
            End
            Begin VB.TextBox Text1 
               Height          =   285
               Index           =   1
               Left            =   1530
               MaxLength       =   2
               TabIndex        =   7
               Top             =   450
               Width           =   465
            End
            Begin VB.TextBox Text1 
               Height          =   285
               Index           =   2
               Left            =   2100
               MaxLength       =   4
               TabIndex        =   8
               Top             =   450
               Width           =   525
            End
            Begin VB.TextBox Text1 
               Height          =   285
               Index           =   3
               Left            =   2700
               MaxLength       =   10
               TabIndex        =   9
               Top             =   450
               Width           =   1155
            End
            Begin VB.TextBox Text1 
               Height          =   285
               Index           =   4
               Left            =   4110
               MaxLength       =   1
               TabIndex        =   10
               Top             =   450
               Width           =   345
            End
            Begin VB.TextBox Text1 
               Height          =   285
               Index           =   5
               Left            =   4560
               MaxLength       =   2
               TabIndex        =   11
               Top             =   450
               Width           =   405
            End
            Begin VB.TextBox Text1 
               Height          =   285
               Index           =   6
               Left            =   5220
               MaxLength       =   4
               TabIndex        =   12
               Top             =   450
               Width           =   585
            End
            Begin VB.Frame Frame6 
               Height          =   555
               Left            =   60
               TabIndex        =   21
               Top             =   720
               Width           =   10755
               Begin VB.Label Label3 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  ForeColor       =   &H000080FF&
                  Height          =   285
                  Index           =   2
                  Left            =   9150
                  TabIndex        =   24
                  Top             =   210
                  Width           =   1395
               End
               Begin VB.Label Label3 
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  ForeColor       =   &H00000000&
                  Height          =   285
                  Index           =   1
                  Left            =   180
                  TabIndex        =   23
                  Top             =   210
                  Width           =   7395
               End
               Begin VB.Label Label3 
                  Caption         =   "Saldo Doc"
                  BeginProperty Font 
                     Name            =   "System"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000040C0&
                  Height          =   285
                  Index           =   0
                  Left            =   7980
                  TabIndex        =   22
                  Top             =   240
                  Width           =   975
               End
            End
            Begin MSMask.MaskEdBox Mbox3 
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "d/MM/yyyy"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   10250
                  SubFormatType   =   3
               EndProperty
               Height          =   255
               Left            =   7440
               TabIndex        =   14
               Top             =   480
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   450
               _Version        =   393216
               MaxLength       =   10
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               Caption         =   "Fecha Doc"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   12
               Left            =   7440
               TabIndex        =   63
               Top             =   240
               Width           =   1125
            End
            Begin VB.Label Label2 
               Caption         =   "Banco"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   8
               Left            =   8805
               TabIndex        =   51
               Top             =   225
               Width           =   585
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               Caption         =   "Importe"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   7
               Left            =   11355
               TabIndex        =   50
               Top             =   225
               Width           =   1005
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               Caption         =   "Mon."
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   6
               Left            =   9630
               TabIndex        =   49
               Top             =   225
               Width           =   465
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               Caption         =   "TD"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   5
               Left            =   4650
               TabIndex        =   48
               Top             =   225
               Width           =   285
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               Caption         =   "P/T"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   4
               Left            =   4125
               TabIndex        =   47
               Top             =   225
               Width           =   315
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               Caption         =   "Numero"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   3
               Left            =   2700
               TabIndex        =   46
               Top             =   225
               Width           =   1155
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               Caption         =   "Serie"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   2
               Left            =   2100
               TabIndex        =   45
               Top             =   225
               Width           =   615
            End
            Begin VB.Label Label2 
               Caption         =   "TD"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   1
               Left            =   1650
               TabIndex        =   44
               Top             =   225
               Width           =   315
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               Caption         =   "Cliente"
               BeginProperty Font 
                  Name            =   "Arial"
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
               TabIndex        =   43
               Top             =   225
               Width           =   1095
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               Caption         =   "Numero"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   9
               Left            =   5925
               TabIndex        =   42
               Top             =   225
               Width           =   1125
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               Caption         =   "Serie"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   10
               Left            =   5250
               TabIndex        =   41
               Top             =   225
               Width           =   615
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               Caption         =   "T. Cambio"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   11
               Left            =   10230
               TabIndex        =   40
               Top             =   225
               Width           =   945
            End
         End
         Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
            Height          =   4275
            Left            =   150
            TabIndex        =   4
            Top             =   300
            Width           =   10845
            _ExtentX        =   19129
            _ExtentY        =   7541
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
      End
   End
   Begin MSComctlLib.StatusBar Panel 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   19
      Top             =   7845
      Width           =   13530
      _ExtentX        =   23865
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
Attribute VB_Name = "FrmPlanillaCobranza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adll As New dllgeneral.dll_general
Dim rsdeta As New ADODB.Recordset
Dim rsdetab As New ADODB.Recordset
Dim tipodoc As String

Private Sub cAcepto_Click()
   Dim J As Double
   If rsdetab.RecordCount > 0 Then
      rsdetab.MoveFirst
      Do Until rsdetab.EOF
         If rsdetab.Fields(0) = "*" Then
            rsdeta.AddNew
            rsdeta.Fields(0) = rsdetab.Fields(1)
            rsdeta.Fields(1) = rsdetab.Fields(2)
            rsdeta.Fields(2) = rsdetab.Fields(3)
            rsdeta.Fields(3) = rsdetab.Fields(4)
            rsdeta.Fields(4) = rsdetab.Fields(5)
            rsdeta.Fields(5) = rsdetab.Fields(6)
            rsdeta.Fields(6) = rsdetab.Fields(7)
            rsdeta.Fields(7) = rsdetab.Fields(8)
            rsdeta.Fields(8) = rsdetab.Fields(9)
            rsdeta.Fields(9) = rsdetab.Fields(10)
            rsdeta.Fields(10) = rsdetab.Fields(11)
            rsdeta.Fields(11) = rsdetab.Fields(12)
            rsdeta.Fields(12) = rsdetab.Fields(13)
            rsdeta.Update
         End If
         rsdetab.MoveNext
      Loop
   End If
   Set rsdetab = Nothing
   Frame8.Visible = False
End Sub

Private Sub cAyuda_Click(Index As Integer)
 Dim SQL As String
   Dim nsaldo As Double
   Dim nestado As Double
  nAyuda = "": nDetalle = "": nsaldo = 0: nestado = 0
  If Index = 0 Then
     If adll.VerificaDatoExistente(VGCNx, "select * from vt_cargo where clientecodigo='" & Text1(0) & "'") = 1 Then
        Dim gfiltra(1, 2) As String
        gfiltra(1, 1) = "Documento": gfiltra(1, 2) = "cargonumdoc"
        FrmAyudaCli.TipoForma = 1
        FrmAyudaCli.BConexion = VGCNx
        FrmAyudaCli.Bdata = "0"
        FrmAyudaCli.BTabla = "vt_cargo A inner join cc_tipodocumento B On a.documentocargo=b.tdocumentocodigo"
        FrmAyudaCli.BCampos = "documentocargo as TD,cargonumdoc as Numero,monedacodigo as Mnd,cargoapeimpape as Total,round((cargoapeimpape-cargoapeimppag),2) as Saldo"
        FrmAyudaCli.BOrden = "Clientecodigo,cargoapefecemi"
        If VGPlanillaAjuste <> 2 Then
           FrmAyudaCli.BCondi = " empresacodigo='" & Ctr_Ayuempresa.xclave & "' and clientecodigo='" & Text1(0) & "' and (Round(cargoapeimpape,2)-Round(isnull(cargoapeimppag,0),2)>0) and b.tdocumentotipo='C' and isnull(cargoapeflgreg,0)<>'1' "
         Else
           FrmAyudaCli.BCondi = " empresacodigo='" & Ctr_Ayuempresa.xclave & "' and clientecodigo='" & Text1(0) & "' and (Round(cargoapeimpape,2)-Round(isnull(cargoapeimppag,0),2)>0) and b.tdocumentotipo='A' and isnull(cargoapeflgreg,0)<>'1' "
         End If
        FrmAyudaCli.BFiltro = gfiltra
        FrmAyudaCli.Show 1
        Text1(1) = nAyuda
        Text1(2) = Left(nDetalle, Text1(2).MaxLength)
        
        If Len(nDetalle) >= Text1(3).MaxLength + Text1(2).MaxLength Then
           Text1(3) = Right(nDetalle, Text1(3).MaxLength)
         Else
          If Len(nDetalle) > 0 Then Text1(3) = Right(nDetalle, Len(nDetalle) - Text1(2).MaxLength)
        End If
        Label3(2).Caption = nsaldo
        Text1(10) = nsaldo
        Text1(4).SetFocus
        
        
     Else
         nAyuda = "": nDetalle = ""
         MsgBox "No existen Documentos...", vbInformation, MsgTitle
         Exit Sub
     End If
   ElseIf Index = 1 Then
    If VGPlanillaAjuste <> 2 Then
       SQL = "select * from cc_tipodocumento where tdocumentotipo='A' and tdocumentoingcobra='1'"
     Else
       SQL = "select * from cc_tipodocumento where tdocumentotipo='C' and tdocumentoingcobra='1'"
     End If
If adll.VerificaDatoExistente(VGCNx, SQL) = 1 Then
        Dim sfiltra(1, 2) As String
        sfiltra(1, 1) = "Documento": sfiltra(1, 2) = "tdocumentocodigo"
        FrmAyudaCli.TipoForma = 1
        FrmAyudaCli.BConexion = VGCNx
        FrmAyudaCli.Bdata = "0"
        FrmAyudaCli.BTabla = "cc_tipodocumento"
        FrmAyudaCli.BCampos = "tdocumentocodigo as Codigo,tdocumentodescripcion as Descripcion"
        FrmAyudaCli.BOrden = "tdocumentocodigo"
        If VGPlanillaAjuste <> 2 Then
           FrmAyudaCli.BCondi = " tdocumentotipo='A' and tdocumentocancela='1'"
        Else
           FrmAyudaCli.BCondi = " tdocumentotipo='C' and tdocumentocancela='1'"
        End If
        FrmAyudaCli.BFiltro = sfiltra
        FrmAyudaCli.Show 1
        Text1(5) = nAyuda
        If Text1(6).Enabled = True Then
          Text1(6).SetFocus
        Else
          Text1(8).SetFocus
        End If
     Else
         nAyuda = "": nDetalle = ""
         MsgBox "No existen Documentos...", vbInformation, MsgTitle
         Exit Sub
     End If
     Call Text1_KeyPress(5, 13)
     Exit Sub
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
        Label3(1) = nDetalle
        Text1(1).SetFocus
   ElseIf Index = 3 Then
    'If adll.VerificaDatoExistente(vgcnx, "select * from gr_banco where bancocodigo='" & Text1(9) & "'") = 1 Then
         If Len(RTrim$(Text1(9))) > 0 Then
            SendKeys "{tab}"
            Exit Sub
         End If
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
         Text1(9) = nAyuda
     ElseIf Index = 4 Then
        If adll.VerificaDatoExistente(VGCNx, "select * from vt_cargo where clientecodigo='" & Text1(0) & "'") = 1 Then
           gfiltra(1, 1) = "Documento": gfiltra(1, 2) = "cargonumdoc"
           nestado = 1
           FrmAyudaCli.TipoForma = 1
           FrmAyudaCli.BConexion = VGCNx
           FrmAyudaCli.Bdata = "0"
           FrmAyudaCli.BTabla = "vt_cargo A inner join cc_tipodocumento B On a.documentocargo=b.tdocumentocodigo"
           FrmAyudaCli.BCampos = "documentocargo as TD,cargonumdoc as Numero,monedacodigo as Mnd,cargoapeimpape as Total,round((cargoapeimpape-cargoapeimppag),2) as Saldo"
           FrmAyudaCli.BOrden = "Clientecodigo,cargoapefecemi"
           SQL = " empresacodigo='" & Ctr_Ayuempresa.xclave & "' and clientecodigo='" & Text1(0) & "' and (Round(cargoapeimpape,2)-Round(isnull(cargoapeimppag,0),2)>0) and b.tdocumentotipo='A' and isnull(cargoapeflgreg,0)<>'1' "
           SQL = SQL & " and documentocargo='" & Text1(5) & "'"
           FrmAyudaCli.BCondi = SQL
           FrmAyudaCli.BFiltro = gfiltra
           FrmAyudaCli.Show 1
          ' Text1(1) = nAyuda
          Text1(6) = Left(nDetalle, Text1(6).MaxLength)
          If Len(Text1(6)) > 0 Then
              If Len(nDetalle) >= Text1(6).MaxLength + Text1(7).MaxLength Then
                 Text1(7) = Right(nDetalle, Text1(7).MaxLength)
               Else
                 Text1(7) = Right(nDetalle, Len(nDetalle) - Text1(6).MaxLength)
              End If
            Else
              Text1(7) = ""
          End If
          Label3(2).Caption = nsaldo
         If nestado = 1 And nsaldo < Text1(10) Then
              Text1(10) = Numero(nsaldo)
           End If
           If nDetalle = "" Then nestado = 0
           Text1(8).SetFocus
         Else
           nAyuda = "": nDetalle = ""
           MsgBox "No existen Documentos...", vbInformation, MsgTitle
           Exit Sub
       End If
    End If
   nAyuda = "": nDetalle = ""
End Sub

Private Sub cCerrar_Click()
  Set rsdetab = Nothing
  Frame8.Visible = False
End Sub

Private Sub cContado_Click()
  Dim rsbusca As New ADODB.Recordset
   
  Frame8.Visible = True
  DoEvents
  Call Carga_Grilla2
  Set rsbusca = VGCNx.Execute("select vt_cargo.clientecodigo,documentocargo,cargonumdoc,monedacodigo,cargoapeimpape,cargoapetipcam from vt_Cargo inner join vt_pedido on vt_pedido.pedidotipofac=vt_cargo.documentocargo  and  vt_pedido.pedidonrofact=vt_cargo.cargonumdoc where cargoapefecemi='" & MBox1 & "' and (cargoapeflgreg<>'1' or cargoapeflgreg is null) and formapagocodigo='01' order by documentocargo")
  If rsbusca.RecordCount > 0 Then
    rsbusca.MoveFirst
    Do Until rsbusca.EOF
        rsdetab.AddNew
        rsdetab.Fields("flag") = "*"
        rsdetab.Fields("Cliente") = rsbusca.Fields("clientecodigo")
        rsdetab.Fields("TD") = rsbusca.Fields("documentocargo")
        rsdetab.Fields("Serie") = Left$(rsbusca.Fields("cargonumdoc"), 4)
        rsdetab.Fields("Numero") = Right$(rsbusca.Fields("cargonumdoc"), 10)
        rsdetab.Fields("P/T") = "T"
        rsdetab.Fields("TDp") = "10"
        rsdetab.Fields("Seriep") = "0000"
        rsdetab.Fields("Numerop") = "0000000000"
        rsdetab.Fields("Moneda") = rsbusca.Fields("monedacodigo")
        rsdetab.Fields("Banco") = ""
        rsdetab.Fields("Importe") = rsbusca.Fields("cargoapeimpape")
        rsdetab.Fields("TCambio") = rsbusca.Fields("cargoapetipcam")
        rsbusca.MoveNext
    Loop
  End If
  rsbusca.Close
  Set rsbusca = Nothing

End Sub


'FIXIT: Declare 'Carga_Grilla2' con un tipo de datos de enlace en tiempo de compilación     FixIT90210ae-R1672-R1B8ZE
Public Sub Carga_Grilla2()
   Set rsdetab = Nothing
    
   Call rsdetab.Fields.Append("flag", adChar, 1)
   Call rsdetab.Fields.Append("Cliente", adVarChar, 11)
   Call rsdetab.Fields.Append("TD", adChar, 2)
   Call rsdetab.Fields.Append("Serie", adChar, 4)
   Call rsdetab.Fields.Append("Numero", adChar, 10)
   Call rsdetab.Fields.Append("P/T", adChar, 1)
   Call rsdetab.Fields.Append("TDp", adChar, 2)
   Call rsdetab.Fields.Append("Seriep", adChar, 4)
   Call rsdetab.Fields.Append("Numerop", adChar, 10)
   Call rsdetab.Fields.Append("Moneda", adChar, 2)
   Call rsdetab.Fields.Append("Banco", adChar, 2)
   Call rsdetab.Fields.Append("Importe", adDouble)
   Call rsdetab.Fields.Append("TCambio", adDouble)
   
   rsdetab.Open
   Set DGrid1.DataSource = rsdetab
   DGrid1.Refresh
   Call ConfigGrid2

End Sub


'FIXIT: Declare 'ConfigGrid2' con un tipo de datos de enlace en tiempo de compilación      FixIT90210ae-R1672-R1B8ZE
Public Sub ConfigGrid2()
    With DGrid1
        .Columns(0).Width = 300
        .Columns(1).Width = 1200
        .Columns(2).Width = 800
        .Columns(3).Width = 500
        .Columns(4).Width = 1300
        .Columns(5).Width = 800
        .Columns(6).Width = 800
        .Columns(7).Width = 800
        .Columns(8).Width = 1300
        .Columns(9).Width = 800
        .Columns(10).Width = 1000
        .Columns(11).Width = 1300
        .Columns(11).NumberFormat = "##,###,###,##0.00"
        .Columns(12).Width = 1000
        .Columns(12).NumberFormat = "##,###,###,##0.00"
    End With
    DGrid1.Refresh
End Sub

Private Sub cmdBotones_Click(Index As Integer)
  Dim acmd As New ADODB.Command
  Dim rb As New ADODB.Recordset
  Dim xnumdoc As String

'FIXIT: Declare 'xabono' and 'xzona' and 'xmone' and 'xcuenta' con un tipo de datos de enlace en tiempo de compilación     FixIT90210ae-R1672-R1B8ZE
  Dim xabono, xzona, xmone, xcuenta, xtipo As String
'FIXIT: Declare 'xnumplan' and 'ximpsol' and 'xtcam' con un tipo de datos de enlace en tiempo de compilación     FixIT90210ae-R1672-R1B8ZE
  Dim xnumplan, ximpsol, xtcam, xnumpag As Double
  Dim nflag As Integer
  Dim mescierre As Boolean
  On Error GoTo nerror
  xabono = ""
  xzona = ""
  xcuenta = ""
  Select Case Index
    Case 0
       Limpiartexto Text1, 0, 11
       Text1(11) = DatoTipoCambio(VGcnxCT, MBox1.Text)
       Text1(0).SetFocus
    Case 1   'GRABAR DATOS
       cmdBotones(1).Enabled = False
        If rsdeta.RecordCount > 0 Then
            Set rb = VGCNx.Execute("select * from cc_tipoplanilla where tplanillacodigo='" & Escadena(Ctr_Ayuda1.xclave) & "'")
            If rb.RecordCount > 0 Then
               xnumplan = Val(RTrim$(rb!tplanillanumerador)) + 1
            Else
               xnumplan = 1
            End If
            rb.Close
            Set rb = Nothing
            nflag = 1
            VGCNx.BeginTrans
            
            VGCNx.Execute "update cc_tipoplanilla " & _
                       " set tplanillanumerador='" & xnumplan & "' " & _
                       " where tplanillacodigo='" & Escadena(Ctr_Ayuda1.xclave) & "'"
        
            rsdeta.MoveLast
            rsdeta.MoveFirst
            Do Until rsdeta.EOF
                xtipo = 0
                Set rb = VGCNx.Execute("select * from cc_tipodocumento where tdocumentocodigo='" & rsdeta.Fields(5) & "'")
                If rb.RecordCount > 0 Then
                   xabono = rb!tdocumentotipo
                   xtipo = IIf(IsNull(rb!tdocumentopermiteaplica), Null, rb!tdocumentopermiteaplica)
                   If rsdeta.Fields(8) = g_TipoSol Then
                      xcuenta = "" & RTrim$(rb!tdocumentocuentasoles)
                   Else
                      xcuenta = "" & RTrim$(rb!tdocumentocuentadolares)
                   End If
                Else
                   xabono = "": xcuenta = "": xtipo = ""
                End If
                rb.Close
                Set rb = Nothing
                
                Set rb = VGCNx.Execute("select * from vt_cargo where empresacodigo='" & Ctr_Ayuempresa.xclave & "' and  documentocargo='" & rsdeta.Fields(1) & "' and cargonumdoc='" & rsdeta.Fields(2) & rsdeta.Fields(3) & "'")
                If rb.RecordCount > 0 Then
                   xzona = rb!zonacodigo
                   xmone = rb!monedacodigo
                   If IsNull(rb!cargoapenumpag) Then
                     xnumpag = 1
                   Else
                     xnumpag = Val(rb!cargoapenumpag)
                   End If
                Else
                   xzona = "01": xmone = g_TipoSol: xnumpag = 1
                End If
                rb.Close
                Set rb = Nothing
                
                ximpsol = CDbl(rsdeta.Fields(10))
                xtcam = rsdeta.Fields(11)
                If rsdeta.Fields(8) <> xmone Then
                   If rsdeta.Fields(8) = g_TipoSol Then
                      xtcam = rsdeta.Fields(11)
                      If rsdeta.Fields(11) = 0 Then xtcam = 1
                      ximpsol = CDbl(rsdeta.Fields(10)) / CDbl(xtcam)
                   Else
                      xtcam = rsdeta.Fields(11)
                      If rsdeta.Fields(11) = 0 Then xtcam = 1
                       ximpsol = CDbl(rsdeta.Fields(10)) * CDbl(xtcam)
                   End If
                End If
                                
                Set acmd.ActiveConnection = VGGeneral
                acmd.CommandType = adCmdStoredProc
                acmd.CommandText = "cc_abonadocumento_pro"
                acmd.CommandTimeout = 0
                acmd.Prepared = True
                With acmd
                    .Parameters("@base") = VGCNx.DefaultDatabase
                    .Parameters("@tipo") = "1"
                    .Parameters("@empresa") = Ctr_Ayuempresa.xclave
                    .Parameters("@documentoabono") = rsdeta.Fields(1)
                    .Parameters("@abononumdoc") = RTrim$(rsdeta.Fields(2) & rsdeta.Fields(3))
                    .Parameters("@abonocannumpag") = xnumpag
                    .Parameters("@zonacodigo") = xzona
                    .Parameters("@tipoplanilla") = Escadena(Ctr_Ayuda1.xclave)
                    .Parameters("@vendedor") = Escadena(Ctr_Ayuda2.xclave)
                    .Parameters("@numplanilla") = Right$("00000000" & RTrim$(CStr(xnumplan)), 6)
                    .Parameters("@fechapla") = MBox1.Text
                    .Parameters("@fechapro") = MBox1.Text
                    .Parameters("@moneda") = xmone
                    .Parameters("@abonocancarabo") = "A"  'xabono
                    .Parameters("@cuenta") = xcuenta
                    .Parameters("@banco") = "" & RTrim$(rsdeta.Fields(9))
                    .Parameters("@tipocam") = CDbl(xtcam)
                    .Parameters("@abonoflpres") = "1"
                    .Parameters("@abonocanimpcan") = CDbl(rsdeta.Fields(10))
                    .Parameters("@abonocanimpsol") = ximpsol
                    .Parameters("@usuario") = g_usuario
                    .Parameters("@fechaact") = Date
                    .Parameters("@forma") = rsdeta.Fields(4)
                    .Parameters("@monedacan") = rsdeta.Fields(8)
                    .Parameters("@abonocantd") = rsdeta.Fields(5)
                    .Parameters("@abonocannro") = RTrim$(rsdeta.Fields(6) & rsdeta.Fields(7))
                    If Len(rsdeta.Fields(12)) > 0 Then
                       .Parameters("@fechacan") = rsdeta.Fields(12)
                     Else
                       .Parameters("@fechacan") = MBox1.Text
                    End If
                    .Parameters("@cliente") = rsdeta.Fields(0)
                    .Parameters("@empresa") = Ctr_Ayuempresa.xclave
                End With
                acmd.Execute
                
                Set acmd = Nothing
                DoEvents
                                
                '**** Actualizamos Saldos de documento pendiente
                If rsdeta.Fields(8) = g_TipoDolar Then
                   If xmone = g_TipoSol Then
                       VGCNx.Execute "Update  vt_cargo Set cargoapeimppag=isnull(cargoapeimppag,0)+" & CDbl(rsdeta.Fields(10) * xtcam) & "," & _
                                   " cargoapenumpag='" & xnumpag + 1 & "'" & _
                                  " Where empresacodigo='" & Ctr_Ayuempresa.xclave & "' and  documentocargo='" & rsdeta.Fields(1) & "' and cargonumdoc='" & RTrim$(rsdeta.Fields(2) & rsdeta.Fields(3)) & "'"
                   Else
                       VGCNx.Execute "Update  vt_cargo Set cargoapeimppag=isnull(cargoapeimppag,0)+" & CDbl(rsdeta.Fields(10)) & "," & _
                                  " cargoapenumpag='" & xnumpag + 1 & "'" & _
                                  " Where empresacodigo='" & Ctr_Ayuempresa.xclave & "' and documentocargo='" & rsdeta.Fields(1) & "' and cargonumdoc='" & RTrim$(rsdeta.Fields(2) & rsdeta.Fields(3)) & "'"
                   End If
                ElseIf rsdeta.Fields(8) = g_TipoSol Then
                   If xmone = g_TipoDolar Then
                       VGCNx.Execute "Update  vt_cargo Set cargoapeimppag=isnull(cargoapeimppag,0)+" & CDbl(rsdeta.Fields(10) / xtcam) & "," & _
                                  " cargoapenumpag='" & xnumpag + 1 & "'" & _
                                  " Where empresacodigo='" & Ctr_Ayuempresa.xclave & "' and documentocargo='" & rsdeta.Fields(1) & "' and cargonumdoc='" & RTrim$(rsdeta.Fields(2) & rsdeta.Fields(3)) & "'"
                   Else
                       VGCNx.Execute "Update  vt_cargo Set cargoapeimppag=isnull(cargoapeimppag,0)+" & CDbl(rsdeta.Fields(10)) & "," & _
                                  " cargoapenumpag='" & xnumpag + 1 & "'" & _
                                  " Where empresacodigo='" & Ctr_Ayuempresa.xclave & "' and documentocargo='" & rsdeta.Fields(1) & "' and cargonumdoc='" & RTrim$(rsdeta.Fields(2) & rsdeta.Fields(3)) & "'"
                   End If
                End If
                
                VGCNx.Execute "Update  vt_cargo " & _
                            " Set cargoapeflgcan= '0'" & _
                            " Where empresacodigo='" & Ctr_Ayuempresa.xclave & "' and documentocargo='" & rsdeta.Fields(5) & "' and cargonumdoc='" & RTrim$(rsdeta.Fields(6) & rsdeta.Fields(7)) & "'"
                                        
                VGCNx.Execute "Update  vt_cargo " & _
                            " Set cargoapeflgcan= '1'," & _
                            "   cargoapefeccan='" & MBox1.Text & "'" & _
                            " Where empresacodigo='" & Ctr_Ayuempresa.xclave & "' and documentocargo='" & rsdeta.Fields(1) & "' and cargonumdoc='" & RTrim$(rsdeta.Fields(2) & rsdeta.Fields(3)) & "'" & _
                            " and cargoapeimpape-isnull(cargoapeimppag,0)= 0 "
                
                                                
                '****Permite Aplicaciones
                
                If Not IsNull(xtipo) Then
                    If xtipo = 1 Then
                            Set rb = VGCNx.Execute("select * from cc_tipodocumento where tdocumentocodigo='" & rsdeta.Fields(1) & "'")
                            If rb.RecordCount > 0 Then
                               xabono = rb!tdocumentotipo
                               If rsdeta.Fields(8) = g_TipoSol Then
                                  xcuenta = "" & RTrim$(rb!tdocumentocuentasoles)
                               Else
                                  xcuenta = "" & RTrim$(rb!tdocumentocuentadolares)
                               End If
                            Else
                               xabono = "": xcuenta = ""
                            End If
                            rb.Close
                            Set rb = Nothing
                            
                            Set rb = VGCNx.Execute("select * from vt_cargo where empresacodigo='" & Ctr_Ayuempresa.xclave & "' and documentocargo='" & rsdeta.Fields(5) & "' and cargonumdoc='" & rsdeta.Fields(6) & rsdeta.Fields(7) & "'")
                            If rb.RecordCount > 0 Then
                               xzona = rb!zonacodigo
                               xmone = rb!monedacodigo
                               If IsNull(rb!cargoapenumpag) Then
                                 xnumpag = 1
                               Else
                                 xnumpag = Val(rb!cargoapenumpag)
                               End If
                            Else
                               xzona = "01": xmone = g_TipoSol: xnumpag = 1
                            End If
                            rb.Close
                            Set rb = Nothing
                                            
'                            Set rb = vgcnx.Execute("select * from cc_tipoplanilla where tplanillacodigo='" & Escadena(Ctr_Ayuda1.xclave) & "'")
'                            If rb.RecordCount > 0 Then
'                               xnumplan = Val(rtrim$(rb!tplanillanumerador)) + 1
'                            Else
'                               xnumplan = 1
'                            End If
'                            rb.Close
'                            Set rb = Nothing
'
'                            vgcnx.Execute "update cc_tipoplanilla " & _
'                                       " set tplanillanumerador='" & xnumplan & "' " & _
'                                       " where tplanillacodigo='" & Escadena(Ctr_Ayuda1.xclave) & "'"
'
                            
                            ximpsol = CDbl(rsdeta.Fields(10))
                            xtcam = rsdeta.Fields(11)
                            If rsdeta.Fields(8) <> xmone Then
                               If rsdeta.Fields(8) = g_TipoSol Then
                                  xtcam = rsdeta.Fields(11)
                                  If rsdeta.Fields(11) = 0 Then xtcam = 1
                                  ximpsol = CDbl(rsdeta.Fields(10)) / CDbl(xtcam)
                               Else
                                  xtcam = rsdeta.Fields(11)
                                  If rsdeta.Fields(11) = 0 Then xtcam = 1
                                   ximpsol = CDbl(rsdeta.Fields(10)) * CDbl(xtcam)
                               End If
                            End If
                                            
                            Set acmd.ActiveConnection = VGGeneral
                            acmd.CommandType = adCmdStoredProc
                            acmd.CommandText = "cc_abonadocumento_pro"
                            acmd.CommandTimeout = 0
                            acmd.Prepared = True
                            With acmd
                                .Parameters("@base") = VGCNx.DefaultDatabase
                                .Parameters("@tipo") = "1"
                                .Parameters("@empresa") = Ctr_Ayuempresa.xclave
                                .Parameters("@documentoabono") = rsdeta.Fields(5)
                                .Parameters("@abononumdoc") = RTrim$(rsdeta.Fields(6) & rsdeta.Fields(7))
                                .Parameters("@abonocannumpag") = xnumpag
                                .Parameters("@zonacodigo") = xzona
                                .Parameters("@tipoplanilla") = Escadena(Ctr_Ayuda1.xclave)
                                .Parameters("@vendedor") = Escadena(Ctr_Ayuda2.xclave)
                                .Parameters("@numplanilla") = Right$("00000000" & RTrim$(CStr(xnumplan)), 6)
                                .Parameters("@fechapla") = MBox1.Text
                                .Parameters("@fechapro") = MBox1.Text
                                .Parameters("@moneda") = xmone
                                .Parameters("@abonocancarabo") = "A"   'xabono
                                .Parameters("@cuenta") = xcuenta
                                .Parameters("@banco") = "" & RTrim$(rsdeta.Fields(9))
                                .Parameters("@tipocam") = CDbl(xtcam)
                                .Parameters("@abonoflpres") = "1"
                                .Parameters("@abonocanimpcan") = CDbl(rsdeta.Fields(10))
                                .Parameters("@abonocanimpsol") = ximpsol
                                .Parameters("@usuario") = g_usuario
                                .Parameters("@fechaact") = Date
                                .Parameters("@forma") = rsdeta.Fields(4)
                                .Parameters("@monedacan") = rsdeta.Fields(8)
                                .Parameters("@abonocantd") = rsdeta.Fields(1)
                                .Parameters("@abonocannro") = RTrim$(rsdeta.Fields(2) & rsdeta.Fields(3))
                                .Parameters("@fechacan") = MBox1.Text
                                .Parameters("@cliente") = rsdeta.Fields(0)
                                .Parameters("@empresa") = Ctr_Ayuempresa.xclave
                            End With
                            acmd.Execute
                            
                            Set acmd = Nothing
                            DoEvents
                                            
                            '**** Actualizamos Saldos de documento pendiente
                            If rsdeta.Fields(8) = g_TipoDolar Then
                               If xmone = g_TipoSol Then
                                       VGCNx.Execute "Update  vt_cargo Set cargoapeimppag=isnull(cargoapeimppag,0)+" & CDbl(rsdeta.Fields(10) * xtcam) & "," & _
                                                " cargoapenumpag='" & xnumpag + 1 & "'" & _
                                               " Where empresacodigo='" & Ctr_Ayuempresa.xclave & "' and documentocargo='" & rsdeta.Fields(5) & "' and cargonumdoc='" & RTrim$(rsdeta.Fields(6) & rsdeta.Fields(7)) & "'"
                               Else
                                        VGCNx.Execute "Update  vt_cargo Set cargoapeimppag=isnull(cargoapeimppag,0)+" & CDbl(rsdeta.Fields(10)) & "," & _
                                                   " cargoapenumpag='" & xnumpag + 1 & "'" & _
                                                   " Where empresacodigo='" & Ctr_Ayuempresa.xclave & "' and documentocargo='" & rsdeta.Fields(5) & "' and cargonumdoc='" & RTrim$(rsdeta.Fields(6) & rsdeta.Fields(7)) & "'"
                                   
                               End If
                            ElseIf rsdeta.Fields(8) = g_TipoSol Then
                               If xmone = g_TipoDolar Then
                                   VGCNx.Execute "Update  vt_cargo Set cargoapeimppag=isnull(cargoapeimppag,0)+" & CDbl(rsdeta.Fields(10) / xtcam) & "," & _
                                              " cargoapenumpag='" & xnumpag + 1 & "'" & _
                                              " Where empresacodigo='" & Ctr_Ayuempresa.xclave & "' and documentocargo='" & rsdeta.Fields(5) & "' and cargonumdoc='" & RTrim$(rsdeta.Fields(6) & rsdeta.Fields(7)) & "'"
                               Else
                                   VGCNx.Execute "Update  vt_cargo Set cargoapeimppag=isnull(cargoapeimppag,0)+" & CDbl(rsdeta.Fields(10)) & "," & _
                                              " cargoapenumpag='" & xnumpag + 1 & "'" & _
                                              " Where empresacodigo='" & Ctr_Ayuempresa.xclave & "' and documentocargo='" & rsdeta.Fields(5) & "' and cargonumdoc='" & RTrim$(rsdeta.Fields(6) & rsdeta.Fields(7)) & "'"
                               End If
                            End If
                    
                            VGCNx.Execute "Update  vt_cargo " & _
                                        " Set cargoapeflgcan= '0'" & _
                                        " Where empresacodigo='" & Ctr_Ayuempresa.xclave & "' and documentocargo='" & rsdeta.Fields(5) & "' and cargonumdoc='" & RTrim$(rsdeta.Fields(6) & rsdeta.Fields(7)) & "'"

                            VGCNx.Execute "Update  vt_cargo " & _
                                        " Set cargoapeflgcan= '1'," & _
                                        "   cargoapefeccan='" & MBox1.Text & "'" & _
                                        " Where empresacodigo='" & Ctr_Ayuempresa.xclave & "' and documentocargo='" & rsdeta.Fields(5) & "' and cargonumdoc='" & RTrim$(rsdeta.Fields(6) & rsdeta.Fields(7)) & "'" & _
                                        " and cargoapeimpape-isnull(cargoapeimppag,0)= 0 "
                            
               
                     End If
                End If
                
                rsdeta.MoveNext
            Loop
            nflag = 0
            VGCNx.CommitTrans
            
       End If
       rsdeta.Close
       Set rsdeta = Nothing
       xnumdoc = Right$("00000000" & RTrim$(CStr(xnumplan)), 6)
       MsgBox "Los datos han sido grabados satisfactoriamente, Planilla Numero  ---> " & xnumdoc & "", vbInformation, MsgTitle
       cmdBotones(1).Enabled = True
       If VGparametros.imprimevoucher = 1 Then Call Imprimir(xnumdoc)
       Call adll.ActivaTab(0, 1, SSTab1)
    Case 2
       If TDBGrid1.Row >= 0 Then
         TDBGrid1.Delete
         TDBGrid1.Update
         TDBGrid1.Refresh
       End If
    Case 4
       
       Call adll.ActivaTab(0, 1, SSTab1)
    Case 11
      If Len(RTrim$(Ctr_Ayuda1.xclave)) = 0 Then
        MsgBox "Falta Ingresar Tipo de Planilla...Verifique!!", vbInformation, MsgTitle
        Exit Sub
      End If
      If Len(RTrim$(Ctr_Ayuda2.xclave)) = 0 Then
        MsgBox "Falta Ingresar Oficina/Vendedor...Verifique!!", vbInformation, MsgTitle
        Exit Sub
      End If
      If adll.VerificaDatoExistente(VGCNx, "select * from cc_tipoplanilla where tplanillacobranza='1' and tplanillacodigo='" & Escadena(Ctr_Ayuda1.xclave) & "' ") = 0 Then
            MsgBox "La planilla no es valida para realizar la cobranza...Verifique!!!", vbInformation, MsgTitle
            Ctr_Ayuda1.SetFocus
            Exit Sub
      End If

      
      Set rsdeta = Nothing
      TDBGrid1.ClearFields
      Set TDBGrid1.DataSource = Nothing
      Call cargar_grilla
       
      Limpiartexto Text1, 0, 11
      Text1(11) = DatoTipoCambio(VGcnxCT, MBox1.Text)
      Call adll.ActivaTab(1, 1, SSTab1)
      Text1(0).SetFocus
    Case 12, 4
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

Private Sub DGrid1_DblClick()
  If rsdetab.RecordCount > 0 Then
    If DGrid1.Columns(0).Text = "*" Then
       DGrid1.Columns(0).Text = ""
    Else
       DGrid1.Columns(0).Text = "*"
    End If
    DGrid1.Update
  End If
End Sub

Private Sub Form_Load()
  MostrarForm Me, "C"
  
  MBox1 = Format(Date, "dd/mm/yyyy")
  Call Ctr_Ayuda1.conexion(VGCNx)
  Call Ctr_Ayuempresa.conexion(VGCNx)
  If VGPlanillaAjuste = 0 Then
     Ctr_Ayuda1.Filtro = "tplanillacobranza='1' and isnull(tplanillaajustes,'0')='0'"
   Else
     Ctr_Ayuda1.Filtro = " isnull(tplanillaajustes,'0')='1'"
  End If
  Call Ctr_Ayuda2.conexion(VGCNx)
    If VGparametros.sistemamultiempresas = True Then
     Lblempresa.Visible = True
     Ctr_Ayuempresa.Visible = True
   Else
     Ctr_Ayuempresa.xclave = VGparametros.empresacodigo
         Lblempresa.Visible = False
     Ctr_Ayuempresa.Visible = False
  End If
  Call adll.ActivaTab(0, 1, SSTab1)
    
End Sub

Public Sub ConfigGrid()
    With TDBGrid1
        .Columns(0).Width = 1200
        .Columns(1).Width = 700
        .Columns(2).Width = 700
        .Columns(3).Width = 1200
        .Columns(4).Width = 700
        .Columns(5).Width = 700
        .Columns(6).Width = 700
        .Columns(7).Width = 1100
        .Columns(8).Width = 700
        .Columns(9).Width = 700
        .Columns(10).Width = 1100
        .Columns(10).NumberFormat = "###,###,##0.00"
        .Columns(11).Width = 900
        .Columns(11).NumberFormat = "#0.0000"
        .Columns(12).Width = 1200
        .Columns(11).NumberFormat = "##/##/####"
        .Refresh
    End With
End Sub

Public Sub cargar_grilla()

   Set rsdeta = Nothing
    
   Call rsdeta.Fields.Append("Cliente", adVarChar, 11)
   Call rsdeta.Fields.Append("TD", adChar, 2)
   Call rsdeta.Fields.Append("Serie", adChar, 4)
   Call rsdeta.Fields.Append("Numero", adChar, 10)
   Call rsdeta.Fields.Append("P/T", adChar, 1)
   Call rsdeta.Fields.Append("TDp", adChar, 2)
   Call rsdeta.Fields.Append("Seriep", adChar, 4)
   Call rsdeta.Fields.Append("Numerop", adChar, 10)
   Call rsdeta.Fields.Append("Moneda", adChar, 2)
   Call rsdeta.Fields.Append("Banco", adChar, 2)
   Call rsdeta.Fields.Append("Importe", adDouble)
   Call rsdeta.Fields.Append("TCambio", adDouble)
   Call rsdeta.Fields.Append("fechadoc", adDate)
   rsdeta.Open
   Set TDBGrid1.DataSource = rsdeta
   TDBGrid1.Refresh
   Call ConfigGrid
End Sub


Private Sub Form_Unload(Cancel As Integer)
'   rsdeta.Close
'   Set rsdeta = Nothing
End Sub
Private Sub MBox1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then
     SendKeys "{tab}"
  End If
  
End Sub

Private Sub MBox3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
     SendKeys "{tab}"
  End If
End Sub

Private Sub Text1_GotFocus(Index As Integer)
 Dim rb As New ADODB.Recordset
 On Error Resume Next
 
  If Index = 4 Then
       Set rb = VGCNx.Execute("select * from vt_cargo where empresacodigo='" & Ctr_Ayuempresa.xclave & "' and  clientecodigo='" & RTrim$(Escadena(Text1(0))) & "' and documentocargo='" & Text1(1) & "' and cargonumdoc='" & RTrim$(Text1(2) & Text1(3)) & "'")
       If rb.RecordCount = 0 Then
           MsgBox "El documento a cancelar del cliente no existe....Verifique!!!", vbInformation, MsgTitle
           rb.Close
           Set rb = Nothing
           Label3(2) = ""
           Text1(3).SetFocus
           Exit Sub
       Else
           If IsNull(rb!cargoapeimppag) Then
                   Label3(2) = DatoMoneda(rb!monedacodigo) & "  - " & Numero(rb!cargoapeimpape)
                   Text1(8) = rb!monedacodigo
                   Text1(10) = Numero(rb!cargoapeimpape)
           Else
                   Label3(2) = DatoMoneda(rb!monedacodigo) & " " & Numero(rb!cargoapeimpape - rb!cargoapeimppag)
                   Text1(8) = rb!monedacodigo
                   Text1(10) = Numero(rb!cargoapeimpape - rb!cargoapeimppag)
            End If
       End If
       rb.Close
    ElseIf Index = 8 Or Index = 9 Then
       Text1(7) = Right$("000000000" & RTrim$(Text1(7)), Text1(7).MaxLength)
       If tipodoc = 1 Then
          Set rb = VGCNx.Execute("select * from vt_cargo where empresacodigo='" & Ctr_Ayuempresa.xclave & "' and clientecodigo='" & RTrim$(Escadena(Text1(0))) & "' and documentocargo='" & Text1(5) & "' and cargonumdoc='" & RTrim$(Text1(6) & Text1(7)) & "'")
          If rb.RecordCount = 0 Then
             MsgBox "El documento de Apliacion no existe....Verifique!!!", vbInformation, MsgTitle
             rb.Close
             Set rb = Nothing
             Text1(7) = ""
             Text1(5).SetFocus
             Exit Sub
          End If
       End If
   End If
   Set rb = Nothing
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
  Dim rb As New ADODB.Recordset
'FIXIT: Declare 'xpago' con un tipo de datos de enlace en tiempo de compilación            FixIT90210ae-R1672-R1B8ZE
  Dim xpago, xcam As Double
  On Error Resume Next
  tipodoc = "0"
  If KeyAscii = 13 Then
          
     If Index = 0 Then
       If Val(Text1(0)) = 0 And Len(RTrim$(Text1(0).Text)) > 0 Then
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
          Label3(1) = Escadena(rb!clientecodigo) & "-" & Escadena(rb!clienterazonsocial)
          Text1(1).SetFocus
       End If
       rb.Close
       Set rb = Nothing
       Exit Sub
     ElseIf Index = 1 Then
       Set rb = VGCNx.Execute("select * from cc_tipodocumento where tdocumentocodigo='" & Text1(Index) & "' and tdocumentotipo='C'")
       If rb.RecordCount = 0 Then
           MsgBox "El documento no es valido....Verifique!!!", vbInformation, MsgTitle
           rb.Close
           Set rb = Nothing
           Exit Sub
       End If
       rb.Close
       Set rb = Nothing
     ElseIf Index = 2 Or Index = 6 Then
        If IsNumeric(Text1(Index)) Then
        Text1(Index) = Right$("000000000" & RTrim$(Text1(Index)), Text1(Index).MaxLength)
        End If
     ElseIf Index = 3 Then
       Text1(Index) = Right$("000000000" & RTrim$(Text1(Index)), Text1(Index).MaxLength)
       Set rb = VGCNx.Execute("select * from vt_cargo where clientecodigo='" & RTrim$(Escadena(Text1(0))) & "' and documentocargo='" & Text1(1) & "' and cargonumdoc='" & RTrim$(Text1(2) & Text1(3)) & "' and cargoapeimpape-isnull(cargoapeimppag,0)>0")
       If rb.RecordCount = 0 Then
           MsgBox "El documento del cliente no existe....Verifique!!!", vbInformation, MsgTitle
           rb.Close
           Set rb = Nothing
           Label3(2) = ""
           Exit Sub
       Else
           If IsNull(rb!cargoapeimppag) Then
                Label3(2) = DatoMoneda(rb!monedacodigo) & "  - " & Numero(rb!cargoapeimpape)
                Text1(8) = rb!monedacodigo
                Text1(10) = Numero(rb!cargoapeimpape)
           Else
                Label3(2) = DatoMoneda(rb!monedacodigo) & " " & Numero(rb!cargoapeimpape - rb!cargoapeimppag)
                Text1(8) = rb!monedacodigo
                Text1(10) = Numero(rb!cargoapeimpape - rb!cargoapeimppag)
            End If
       End If
       rb.Close
       Set rb = Nothing
     ElseIf Index = 4 Then
'FIXIT: Reemplazar la función 'UCase' con la función 'UCase$'.                             FixIT90210ae-R9757-R1B8ZE
         Text1(Index) = UCase$(Text1(Index))
         If Not Text1(Index) Like "[PT]" Then
            MsgBox "Debe ser P ó T...Verifique!!!", vbInformation, MsgTitle
            Exit Sub
         End If
     ElseIf Index = 5 Then
        Tipodocu.numeauto = "0"
        Tipodocu.numerador = ""
       If VGPlanillaAjuste <> 2 Then
           Set rb = VGCNx.Execute("select * from cc_tipodocumento where tdocumentocodigo='" & RTrim$(Escadena(Text1(Index))) & "' and tdocumentotipo<>'C' and tdocumentocancela='1'")
        Else
           Set rb = VGCNx.Execute("select * from cc_tipodocumento where tdocumentocodigo='" & RTrim$(Escadena(Text1(Index))) & "' and tdocumentotipo<>'A' and tdocumentocancela='1'")
       End If
       If rb.RecordCount = 0 Then
           MsgBox "El documento no existe y/o no esta habilitado....Verifique!!!", vbInformation, MsgTitle
           rb.Close
           Set rb = Nothing
           Exit Sub
           'Text1(9).Enabled = False
       Else
           If Escadena(rb!tdocumentovalidabanco) = "1" Then
              Text1(9).Enabled = True
           Else
               Text1(9).Enabled = False
           End If
            Tipodocu.numeauto = IIf(rb!tdocumentonumeauto = "1", "1", "0")
            Tipodocu.numerador = IIf(rb!tdocumentonumeauto = "1", RTrim$(rb!tdocumentonumerador), "1")
            
            If Tipodocu.numeauto = "1" Then
               Text1(6) = Right$("0000000", Text1(6).MaxLength)
               Text1(7) = Right$("000000000" & RTrim$(Tipodocu.numerador), Text1(7).MaxLength)
            End If
       End If
       rb.Close
       Set rb = Nothing
     ElseIf Index = 7 Then
       Text1(Index) = Right$("000000000" & RTrim$(Text1(Index)), Text1(Index).MaxLength)
       Set rb = VGCNx.Execute("select * from cc_tipodocumento where tdocumentocodigo='" & Text1(5) & "' and tdocumentotipo='A'")
       If rb.RecordCount = 0 Then
           MsgBox "El documento no es valido....Verifique!!!", vbInformation, MsgTitle
           rb.Close
           Set rb = Nothing
           Exit Sub
         Else
           tipodoc = rb!tdocumentopermiteaplica
       End If
       rb.Close
       Set rb = Nothing
     ElseIf Index = 8 Then
       Set rb = VGCNx.Execute("select * from gr_moneda where monedacodigo='" & RTrim$(Escadena(Text1(Index))) & "'")
       If rb.RecordCount = 0 Then
           MsgBox "La moneda no existe....Verifique!!!", vbInformation, MsgTitle
           rb.Close
           Set rb = Nothing
           Exit Sub
       End If
       rb.Close
       Set rb = Nothing
       
        Set rb = VGCNx.Execute("select * from vt_cargo where documentocargo='" & Text1(1) & "' and cargonumdoc='" & RTrim$(Text1(2) & Text1(3)) & "' and clientecodigo='" & Text1(0) & "'")
        If rb.RecordCount > 0 Then
            If rb!monedacodigo <> Text1(8) Then
              If Text1(8) = g_TipoSol Then
                If IsNull(rb!cargoapeimppag) Then
                   Text1(10) = Numero(rb!cargoapeimpape * CDbl(Text1(11)))
                Else
                   Text1(10) = Numero((rb!cargoapeimpape - rb!cargoapeimppag) * CDbl(Text1(11)))
                End If
              Else
                If IsNull(rb!cargoapeimppag) Then
                  Text1(10) = Numero(rb!cargoapeimpape / CDbl(Text1(11)))
                Else
                   Text1(10) = Numero((rb!cargoapeimpape - rb!cargoapeimppag) / CDbl(Text1(11)))
                End If
              End If
            End If
       Else
            MsgBox "Verifique,no corresponde el documento al cliente!!", vbInformation, MsgTitle
            rb.Close
            Set rb = Nothing
            Text1(0).SetFocus
            Exit Sub
       End If
       rb.Close
       Set rb = Nothing
       
    ElseIf Index = 9 And Text1(9).Enabled = True Then
            Set rb = VGCNx.Execute("select * from gr_banco where bancocodigo='" & RTrim$(Escadena(Text1(Index))) & "'")
            If rb.RecordCount = 0 Then
                MsgBox "Codigo de Banco no existe....Verifique!!!", vbInformation, MsgTitle
                Text1(9) = ""
            End If
            rb.Close
            Set rb = Nothing
    ElseIf Index = 10 Then
        Text1(Index) = Format(RTrim$(Text1(Index)), "##,###,##0.00")
     
        Set rb = VGCNx.Execute("select * from vt_cargo where EMPRESACODIGO='" & Ctr_Ayuempresa.xclave & "' and clientecodigo='" & RTrim$(Escadena(Text1(0))) & "' and documentocargo='" & Text1(1) & "' and cargonumdoc='" & RTrim$(Text1(2) & Text1(3)) & "'")
        If rb.RecordCount > 0 Then
            If rb!monedacodigo = Text1(8) Then
               If IsNull(rb!cargoapeimppag) Then
                  xpago = rb!cargoapeimpape
               Else
                  xpago = rb!cargoapeimpape - rb!cargoapeimppag
               End If
            Else
              xcam = Text1(11)
              If xcam = 0 Then xcam = 1
              ' monedaprov <> monedapago
              'monedaprov='01' : monedapago='02'
              
              If rb!monedacodigo = g_TipoSol Then
                    xpago = (rb!cargoapeimpape - ESNULO(rb!cargoapeimppag, 0)) / xcam
                 Else
                  'monedaprov='02' : monedapago='01'
                        xpago = (rb!cargoapeimpape - ESNULO(rb!cargoapeimppag, 0)) * xcam
              End If
            End If
            Text1(4) = UCase$(Text1(4))
            If Text1(4) = "T" Then
               If CDbl(Round(xpago, 2)) <> Round(CDbl(Text1(Index)), 2) Then
                  MsgBox "El importe Total debe ser : " & Chr(13) & Chr(10) & Numero(xpago), vbInformation, MsgTitle
                  rb.Close
                  Set rb = Nothing
                  Exit Sub
               End If
            ElseIf Text1(4) = "P" Then
               If CDbl(Round(xpago, 2)) <= Round(CDbl(Text1(Index)), 2) Then
                  MsgBox "El importe Total debe ser menor a : " & Chr(13) & Chr(10) & Numero(xpago), vbInformation, MsgTitle
                  rb.Close
                  Set rb = Nothing
                  Exit Sub
               End If
            End If
        End If
        
        rb.Close
        Set rb = Nothing
                
        rsdeta.AddNew
        rsdeta.Fields(0) = Escadena(Text1(0))
        rsdeta.Fields(1) = Escadena(Text1(1))
        rsdeta.Fields(2) = Escadena(Text1(2))
        rsdeta.Fields(3) = Escadena(Text1(3))
        rsdeta.Fields(4) = Escadena(Text1(4))
        rsdeta.Fields(5) = Escadena(Text1(5))
        rsdeta.Fields(6) = Escadena(Text1(6))
        rsdeta.Fields(7) = Escadena(Text1(7))
        rsdeta.Fields(8) = Escadena(Text1(8))
        rsdeta.Fields(9) = Escadena(Text1(9))
        rsdeta.Fields(10) = Numero(IIf(IsNull(Text1(10)) Or Len(RTrim$(Text1(10))) = 0, 0, Text1(10)))
        rsdeta.Fields(11) = IIf(IsNull(Text1(11)) Or Len(RTrim$(Text1(11))) = 0, 0, Text1(11))
        rsdeta.Fields(12) = Mbox3.Text
        rsdeta.Update
        
 '       If Tipodocu.numeauto = "1" Then
 '          VGCNx.Execute "Update cc_tipodocumento " & _
 '                      " Set tdocumentonumerador='" & Val(Text1(7)) + 1 & "'" & _
 '                      " Where tdocumentocodigo='" & Text1(5).Text & "'"
 '
 '       End If
                
        Limpiartexto Text1, 0, 11
        Text1(11) = Format(DatoTipoCambio(VGcnxCT, MBox1.Text), "#0.0000")
        Text1(0).SetFocus
        Exit Sub
    
    ElseIf Index = 11 Then
        Text1(Index) = Format(RTrim$(Text1(Index)), "#0.0000")
     End If
     If Tipodocu.numeauto = "1" And Index = 5 Then
          Set rb = VGCNx.Execute("select * from cc_tipodocumento where tdocumentocodigo='" & RTrim$(Escadena(Text1(Index))) & "' and tdocumentotipo<>'C' and tdocumentopermiteaplica='1'")
          If rb.RecordCount = 0 Then
            rb.Close
            Text1(6).Enabled = False
            Text1(7).Enabled = False
            Text1(8).SetFocus
            Set rb = Nothing
            Exit Sub
          Else
            Text1(6).Enabled = True
            Text1(7).Enabled = True
            Text1(6).SetFocus
            Set rb = Nothing
            Exit Sub
          End If
     Else
        Text1(6).Enabled = True
        Text1(7).Enabled = True
        SendKeys "{tab}"
     End If
     
  End If
  Set rb = Nothing
End Sub

Private Sub Imprimir(nn As String)
Dim Aparam(4) As Variant
Dim aform() As Variant
Aparam(0) = VGCNx.DefaultDatabase
Aparam(1) = Ctr_Ayuempresa.xclave
Aparam(2) = Ctr_Ayuda1.xclave
Aparam(3) = nn
Call ImpresionRptProc("cc_ImpprimePlanillaCobranza.rpt", aform, Aparam, Empty, "Planilla de apliacaciones")
End Sub

