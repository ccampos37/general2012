VERSION 5.00
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Begin VB.Form FrmTraslado 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Traslado entre Almacenes"
   ClientHeight    =   5835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   11910
   Begin VB.Frame Frame5 
      Height          =   675
      Left            =   60
      TabIndex        =   23
      Top             =   1200
      Width           =   11775
      Begin VB.TextBox txtcanti 
         Height          =   285
         Index           =   0
         Left            =   8460
         TabIndex        =   5
         Top             =   210
         Width           =   1215
      End
      Begin VB.TextBox txtcanti 
         Height          =   285
         Index           =   1
         Left            =   10590
         TabIndex        =   6
         Top             =   210
         Width           =   1005
      End
      Begin VB.CommandButton cAyuda 
         Caption         =   "..."
         Height          =   285
         Index           =   3
         Left            =   2430
         TabIndex        =   4
         Top             =   240
         Width           =   285
      End
      Begin MSMask.MaskEdBox MBox2 
         Height          =   285
         Index           =   1
         Left            =   1140
         TabIndex        =   3
         Top             =   240
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   8
         PromptChar      =   "_"
      End
      Begin VB.Label Label2 
         Caption         =   "Cantidad"
         Height          =   285
         Index           =   1
         Left            =   9810
         TabIndex        =   27
         Top             =   270
         Width           =   1155
      End
      Begin VB.Label Label2 
         Caption         =   "Cant. Ref."
         Height          =   285
         Index           =   2
         Left            =   7500
         TabIndex        =   26
         Top             =   270
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Producto"
         Height          =   285
         Index           =   0
         Left            =   270
         TabIndex        =   25
         Top             =   240
         Width           =   1155
      End
      Begin VB.Label Label5 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2760
         TabIndex        =   24
         Top             =   240
         Width           =   4455
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Nota"
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
      Height          =   975
      Left            =   90
      TabIndex        =   20
      Top             =   4410
      Width           =   2835
      Begin VB.Label Label4 
         Caption         =   "Eliminar Item"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Index           =   3
         Left            =   1500
         TabIndex        =   22
         Top             =   390
         Width           =   1095
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   150
         Picture         =   "FrmTraslado.frx":0000
         Top             =   330
         Width           =   480
      End
      Begin VB.Label Label6 
         Caption         =   "[DEL]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   810
         TabIndex        =   21
         Top             =   390
         Width           =   495
      End
   End
   Begin VB.Frame Frame4 
      Height          =   930
      Left            =   5490
      TabIndex        =   10
      Top             =   4800
      Width           =   2880
      Begin VB.CommandButton cmdBotones 
         Caption         =   "&Nuevo"
         Height          =   690
         Index           =   0
         Left            =   1950
         Picture         =   "FrmTraslado.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   180
         Width           =   870
      End
      Begin VB.CommandButton cmdBotones 
         Caption         =   "&Acepta"
         Height          =   690
         Index           =   11
         Left            =   90
         Picture         =   "FrmTraslado.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   180
         Width           =   870
      End
      Begin VB.CommandButton cmdBotones 
         Caption         =   "&Cancelar"
         Height          =   690
         Index           =   12
         Left            =   1050
         Picture         =   "FrmTraslado.frx":0CC6
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   180
         Width           =   825
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1185
      Left            =   60
      TabIndex        =   9
      Top             =   30
      Width           =   11775
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuda1 
         Height          =   375
         Left            =   1800
         TabIndex        =   0
         Top             =   360
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   661
         XcodMaxLongitud =   0
         xcodwith        =   100
         NomTabla        =   "tabalm"
         TituloAyuda     =   "Almacenes"
         ListaCampos     =   "TAALMA(1),TADESCRI(1)"
         XcodCampo       =   "TAALMA"
         XListCampo      =   "TADESCRI"
         ListaCamposDescrip=   "Codigo,Descripcion"
         ListaCamposText =   "TAALMA,TADESCRI"
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuda2 
         Height          =   375
         Left            =   7650
         TabIndex        =   1
         Top             =   390
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   661
         XcodMaxLongitud =   0
         xcodwith        =   100
         NomTabla        =   "tabalm"
         TituloAyuda     =   "Almacenes"
         ListaCampos     =   "TAALMA(1),TADESCRI(1)"
         XcodCampo       =   "TAALMA"
         XListCampo      =   "TADESCRI"
         ListaCamposDescrip=   "Codigo,Descripcion"
         ListaCamposText =   "TAALMA,TADESCRI"
      End
      Begin MSMask.MaskEdBox MBox 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   3
         EndProperty
         Height          =   285
         Index           =   10
         Left            =   1770
         TabIndex        =   2
         Top             =   780
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   503
         _Version        =   393216
         ClipMode        =   1
         AllowPrompt     =   -1  'True
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha "
         Height          =   285
         Index           =   4
         Left            =   270
         TabIndex        =   13
         Top             =   810
         Width           =   825
      End
      Begin VB.Label Label1 
         Caption         =   "Almacen Destino"
         Height          =   195
         Index           =   1
         Left            =   6120
         TabIndex        =   12
         Top             =   420
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Almacen Origen"
         Height          =   195
         Index           =   0
         Left            =   300
         TabIndex        =   11
         Top             =   360
         Width           =   1575
      End
   End
   Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
      Height          =   2205
      Left            =   60
      TabIndex        =   19
      Top             =   1920
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   3889
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
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=9.75,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=System"
      PrintInfos(0).PageFooterFont=   "Size=9.75,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=System"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowUpdate     =   0   'False
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      MultipleLines   =   0
      CellTipsWidth   =   0
      InsertMode      =   0   'False
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
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=18,.bold=0,.fontsize=825,.italic=0"
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
      _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1,.bgcolor=&HC0C0C0&"
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
   Begin VB.Frame Frame2 
      Height          =   555
      Left            =   4620
      TabIndex        =   14
      Top             =   4110
      Width           =   5985
      Begin VB.Label Label4 
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
         Height          =   285
         Index           =   1
         Left            =   4560
         TabIndex        =   18
         Top             =   180
         Width           =   1245
      End
      Begin VB.Label Label4 
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
         Height          =   285
         Index           =   0
         Left            =   1320
         TabIndex        =   17
         Top             =   180
         Width           =   1245
      End
      Begin VB.Label Label3 
         Caption         =   "Nota Ingreso"
         Height          =   195
         Index           =   1
         Left            =   3390
         TabIndex        =   16
         Top             =   240
         Width           =   945
      End
      Begin VB.Label Label3 
         Caption         =   "Nota Salida"
         Height          =   165
         Index           =   0
         Left            =   210
         TabIndex        =   15
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "FrmTraslado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dllgeneral As New dllgeneral.dll_general
Dim rsdeta As New ADODB.Recordset
Dim flag As Integer

Public Function CargaGrilla()
   Set rsdeta = Nothing
   Call rsdeta.Fields.Append("Item", adInteger)
   Call rsdeta.Fields.Append("Codigo", adChar, 20)
   Call rsdeta.Fields.Append("Descripcion", adChar, 100)
   Call rsdeta.Fields.Append("UM", adChar, 3)
   Call rsdeta.Fields.Append("Cant", adDouble)
   Call rsdeta.Fields.Append("CantRef", adDouble)
   
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
      .Columns(2).Width = 4500
      .Columns(2).Caption = "Descripcion"
      .Columns(3).Width = 800
      .Columns(3).Caption = "UM"
      .Columns(4).Width = 1200
      .Columns(4).Caption = "Cant"
      .Columns(4).NumberFormat = "##,###,##0.00"
      .Columns(5).Width = 1200
      .Columns(5).Caption = "Cant.Ref"
      .Columns(5).NumberFormat = "##,###,##0.00"
   End With
   TDBGrid1.Refresh
End Function

Private Sub cAyuda_Click(Index As Integer)
    If Index = 3 Then
        If Len(Label5) > 0 Then
          SendKeys "{tab}"
          Exit Sub
        End If
        Dim sfiltra(1 To 2, 1 To 2) As String
        sfiltra(1, 1) = "Codigo": sfiltra(1, 2) = "acodigo"
        sfiltra(2, 1) = "Descripcion": sfiltra(2, 2) = "adescri"
        FrmAyudaPedidos.TipoForma = 1
        FrmAyudaPedidos.BConexion = VGCNx
        FrmAyudaPedidos.BTabla = "[" & VGCNx.DefaultDatabase & "].dbo.maeart inner join [" & _
                            VGCNx.DefaultDatabase & "].dbo.stkart " & _
                            " ON acodigo=stcodigo"
        FrmAyudaPedidos.Bdata = "4"
        FrmAyudaPedidos.Bdato = Escadena(Trim(MBox2(1).ClipText))
        FrmAyudaPedidos.BCampos = "acodigo as Codigo,adescri as Descripcion,stskdis as stock"
        FrmAyudaPedidos.BOrden = "adescri"
        FrmAyudaPedidos.BCondi = "stalma='" & Ctr_Ayuda1.xclave & "' and stskdis>0"
        FrmAyudaPedidos.BFiltro = sfiltra
        FrmAyudaPedidos.Show 1
        MBox2(1) = Escadena(nAyuda):   Label5 = Escadena(nDetalle)
        Call MBox2_KeyDown(0, 13, 0)
        txtcanti(0).SetFocus
     End If
End Sub

Private Sub cmdBotones_Click(Index As Integer)
  Select Case Index
    Case 0
       MBox2(1) = ""
       txtcanti(0) = "": txtcanti(1) = "": Label5 = ""
       Call CargaGrilla
       Ctr_Ayuda1.SetFocus
    Case 11
      If rsdeta.RecordCount > 0 Then
        GrabarData
      Else
        MsgBox "Debe ingresar productos...verifique!!!", vbInformation, "AVISO"
        Exit Sub
      End If
    Case Else
      Set rsdeta = Nothing
      Unload Me
  End Select
End Sub

Private Sub Ctr_Ayuda3_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)

End Sub

Private Sub Form_Load()
    MostrarFormVentas Me, "C2"
    
    MBox(10) = Format(Date, "DD/MM/YYYY")
    Call Ctr_Ayuda1.conexion(VGCNx)
    Call Ctr_Ayuda2.conexion(VGCNx)
    Call CargaGrilla
End Sub


Public Function GrabarData() As Integer
    Dim J As Integer
    Dim nsql As String
    Dim ltipo As String
    Dim lzona As String
    Dim xserie As String * 3
    Dim xfactu As Double  'String * 8
    Dim xtipofac As String * 2
    Dim ndato As String
    
    Dim acmd As New ADODB.Command
    Dim asql As New ADODB.Recordset
    Dim arbusca As New ADODB.Recordset
    Dim wCabe(40)
    
    GrabarData = 0
    
    '******** CABECERA DE MOVIMIENTO *****************
    For J = 1 To 29
        wCabe(J) = ""
    Next J
    Label4(0) = "": Label4(1) = ""
    
    Set asql = VGCNx.Execute("select * from  num_documentos where ctncodigo='TR'")
    If asql.RecordCount > 0 Then
        ndato = Right("000000000" & Trim(CStr(asql!ctnnumero)), 11)                    'nro pedido"
    Else
       MsgBox " No existe documentos de transacciones...Verifique!!", vbInformation, "AVISO"
       asql.Close
       Set asql = Nothing
       Exit Function
    End If
    asql.Close
    Set asql = Nothing

    VGCNx.Execute "update num_documentos " & _
                    " set ctnnumero=ctnnumero+1 " & _
                    " where ctncodigo='TR'"

    
    For J = 1 To 2
        wCabe(1) = g_ptoventa                        'Pto Venta
        Set asql = Nothing
        If J = 1 Then
            ' de Almacen origen
           Set asql = VGCNx.Execute("select * from tabalm where taalma='" & Ctr_Ayuda1.xclave & "'")
           If asql.RecordCount > 0 Then
               wCabe(2) = Right("0000000000000" & Trim(CStr(asql!tanumsal)), 11)                       'nro pedido"
           End If
           asql.Close
           Set asql = Nothing
           VGCNx.Execute "update tabalm " & _
                           " set tanumsal=tanumsal+1 " & _
                           " where taalma='" & Ctr_Ayuda1.xclave & "'"
                           
           Label4(0) = wCabe(2)
        Else
            ' al almacen destino
           Set asql = VGCNx.Execute("select * from tabalm where taalma='" & Ctr_Ayuda2.xclave & "'")
           If asql.RecordCount > 0 Then
               wCabe(2) = Right("0000000000000" & Trim(CStr(asql!tanument)), 11)                       'nro pedido"
           End If
           asql.Close
           Set asql = Nothing
           VGCNx.Execute "update tabalm " & _
                           " set tanument=tanument+1 " & _
                           " where taalma='" & Ctr_Ayuda2.xclave & "'"
           Label4(1) = wCabe(2)
        End If
        wCabe(3) = ndato                      'nro factura
        wCabe(4) = "TR"                      'nro boleta
        wCabe(5) = ""                      'nro guia
        wCabe(6) = 0                       'dscto gral
        wCabe(7) = 0                       'dscto promocional
        wCabe(8) = 0                       'dscto especial
        wCabe(9) = g_tiposol               'moneda
        wCabe(10) = 0                      'tipo de cambio
        wCabe(11) = 0                      'lista de precios
        wCabe(12) = ""                'mensajes
        wCabe(13) = ""                     'modo de venta
        wCabe(14) = MBox(10)               'fecha de atencion
        wCabe(15) = "00"                   'forma de pago
        wCabe(16) = ""                     'cliente
        wCabe(17) = ""                     'vendedor
        wCabe(18) = 0                      'comision
        If J = 1 Then
           wCabe(19) = Ctr_Ayuda1.xclave           'almacen
        Else
           wCabe(19) = Ctr_Ayuda2.xclave           'almacen
        End If
        wCabe(20) = 0                     'otros gastos
        wCabe(21) = 0                     'nota pedido
        wCabe(22) = 0                     'orden de compra
        wCabe(23) = 0                     'autorizacion
        wCabe(24) = 0                     'dias pago
        wCabe(25) = 0                     'Total Cantidad
        wCabe(26) = 0                     'Total Bruto
        wCabe(27) = 0                     'total fletes --T.D.
        wCabe(28) = 0                     'Total Igv
        wCabe(29) = 0         'Neto a Facturar
        wCabe(30) = ""             'entrega pedido
        wCabe(31) = ""                    'nombre cliente
        wCabe(32) = ""                    'direccion
        wCabe(33) = ""                    'ruc
        wCabe(34) = MBox(10)                           'fechafactura
        wCabe(35) = 0                     'Total Descuentos Globales
        wCabe(36) = 0                    'Total Descuentos Cliente
        wCabe(37) = 0                  'Total Descuentos Oficina
        wCabe(38) = 0                       'Total Descuentos Item
        wCabe(39) = 0                      'Total Descuentos Linea
        wCabe(40) = 0                      'Total Descuentos x Promocion
        
        Set acmd.ActiveConnection = VGgeneral
        acmd.CommandType = adCmdStoredProc
        acmd.CommandText = "vt_ingresoalma_pro"
        acmd.CommandTimeout = 0
        acmd.Prepared = True
        With acmd
            .Parameters("@base") = VGCNx.DefaultDatabase
            .Parameters("@tabla") = "movalmcab"
            If J = 1 Then
              .Parameters("@tipo") = "2"
            Else
              .Parameters("@tipo") = "3"
            End If
            .Parameters("@puntovta") = wCabe(1)
            .Parameters("@numero") = wCabe(2)
            .Parameters("@factura") = wCabe(3)
            .Parameters("@boleta") = wCabe(4)
            .Parameters("@guia") = wCabe(5)
            .Parameters("@dsctoglobal") = wCabe(6)
            .Parameters("@dsctoppago") = wCabe(7)
            .Parameters("@dsctovtaofi") = wCabe(8)
            .Parameters("@moneda") = IIf(wCabe(9) = g_tiposol, "S", "D")
            .Parameters("@tipocambio") = wCabe(10)
            .Parameters("@listaprecio") = wCabe(11)
            .Parameters("@mensaje") = wCabe(12)
            .Parameters("@modoventa") = wCabe(13)
            .Parameters("@fecha") = wCabe(14)
            .Parameters("@formapago") = wCabe(15)
            .Parameters("@cliente") = wCabe(16)
            .Parameters("@vendedor") = wCabe(17)
            .Parameters("@porcomision") = wCabe(18)
            .Parameters("@almacen") = wCabe(19)
            .Parameters("@totalotros") = wCabe(20)
            .Parameters("@notaped") = wCabe(21)
            .Parameters("@ordencompra") = wCabe(22)
            .Parameters("@autoriza") = wCabe(23)
            .Parameters("@diaspago") = wCabe(24)
            .Parameters("@totalitem") = wCabe(25)
            .Parameters("@totalbruto") = wCabe(26)
            .Parameters("@totalflete") = wCabe(27)
            .Parameters("@totalimpuesto") = wCabe(28)
            .Parameters("@totalneto") = wCabe(29)
            .Parameters("@usuario") = g_usuario
            .Parameters("@fechaactual") = Date
            .Parameters("@totaldsctoxlinea") = wCabe(39)
            .Parameters("@montodsctoppago") = 0
            .Parameters("@entregapedido") = wCabe(30)
            .Parameters("@razon") = wCabe(31)
            .Parameters("@direccion") = wCabe(32)
            .Parameters("@ruc") = wCabe(33)
            .Parameters("@fechafactura") = wCabe(34)
            .Parameters("@TDGlobal") = wCabe(35)
            .Parameters("@TDCliente") = wCabe(36)
            .Parameters("@TDOficina") = wCabe(37)
            .Parameters("@TDItem") = wCabe(38)
            .Parameters("@TDPromo") = wCabe(40)
            .Parameters("@empresa") = VGParametros.empresacodigo
        End With
        acmd.Execute
        Set acmd = Nothing
        DoEvents
          
       '** Actualizamos detalle
    
        If rsdeta.RecordCount > 0 Then
            rsdeta.MoveFirst
            Do Until rsdeta.EOF
                If J = 2 Then
                   If dllgeneral.VerificaDatoExistente(VGCNx, "select * from stkart where stalma='" & Ctr_Ayuda2.xclave & "' and stcodigo='" & Trim(rsdeta.Fields(1)) & "'") = 0 Then
                       VGCNx.Execute "insert into stkart " & _
                                       "(stalma,stcodigo,stskdis)" & _
                                       " Values ('" & Ctr_Ayuda2.xclave & "','" & Trim(rsdeta.Fields(1)) & "',0)"
                   End If
                End If

                Set acmd.ActiveConnection = VGgeneral
                acmd.CommandType = adCmdStoredProc
                acmd.CommandTimeout = 0
                acmd.CommandText = "vt_ingresodetallealma_pro"
                acmd.Prepared = True
                With acmd
                    .Parameters("@base") = VGCNx.DefaultDatabase
                    .Parameters("@tabla") = "movalmdet" ' nsql
                    If J = 1 Then
                      .Parameters("@tipo") = "2"
                    Else
                      .Parameters("@tipo") = "3"
                    End If
                    .Parameters("@item") = 1
                    .Parameters("@numero") = wCabe(2)
                    .Parameters("@producto") = Trim(rsdeta.Fields(1))   'Trim(MBox2(1).Text)
                    .Parameters("@unidad") = ""
                    .Parameters("@cantidad") = Trim(rsdeta.Fields(4))   'Trim(txtcanti(1).Text)
                    .Parameters("@preciopacto") = 0
                    .Parameters("@dsctoxitem") = 0
                    .Parameters("@importebruto") = 0
                    .Parameters("@porcomision") = 0
                    .Parameters("@mdsctoitem") = 0
                    .Parameters("@mdsctoxlinea") = 0
                    .Parameters("@mdsctoxprom") = 0
                    .Parameters("@mimpor") = 0
                    .Parameters("@unidadref") = Trim(rsdeta.Fields(5))   'rtxtcanti(1)
                     If J = 1 Then
                       .Parameters("@almacen") = Trim(Ctr_Ayuda1.xclave)
                     Else
                       .Parameters("@almacen") = Trim(Ctr_Ayuda2.xclave)
                     End If
                End With
                acmd.Execute
                Set acmd = Nothing
                            
                Set acmd.ActiveConnection = VGgeneral
                acmd.CommandType = adCmdStoredProc
                acmd.CommandTimeout = 0
                acmd.CommandText = "vt_actualizoalma_pro"
                acmd.Prepared = True
                With acmd
                    .Parameters("@basedes") = CStr(VGCNx.DefaultDatabase)
                    .Parameters("@almacen") = wCabe(19)
                    If J = 1 Then
                      .Parameters("@tipo") = "1"
                    Else
                      .Parameters("@tipo") = "2"
                    End If
                    .Parameters("@articulo") = Trim(rsdeta.Fields(1))   'Trim(MBox2(1).Text)
                    .Parameters("@cantidad") = Trim(rsdeta.Fields(4))   'txtcanti(1)
                End With
                acmd.Execute
                Set acmd = Nothing
                rsdeta.MoveNext
          Loop
       End If
    Next
    GrabarData = 1
    MsgBox "Traslado de almacen satisfactorio...!!", vbInformation, "AVISO"
    
   ' Set rsdeta = Nothing
'vererror:
'   If Err Then
'      MsgBox "Comunicarse con Sistemas ...!!" & Chr(13) & Chr(10) & VGcnx.Errors(0).Number & "-" & VGcnx.Errors(0).Description
'      Exit Function
'   End If

End Function

Private Sub MBox_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then
    MBox2(1).SetFocus
  End If
End Sub

Private Sub MBox2_Change(Index As Integer)
  If Len(Trim(MBox2(1).ClipText)) = 0 Then
    Label5 = ""
  End If
End Sub

Private Sub MBox2_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  Dim nsql As String
  Dim rabusca As New ADODB.Recordset
  
  If KeyCode = 13 Then
     If dllgeneral.ValidaCadena(Trim(MBox2(1).ClipText), "N") = False Then
        MBox2(1).MaxLength = 64
        Call cAyuda_Click(3)
        MBox2(1).MaxLength = 8
        SendKeys "{tab}"
     Else
        MBox2(1).MaxLength = 8
        nsql = "select * from maeart inner join stkart on acodigo=stcodigo  where stcodigo='" & MBox2(1).ClipText & "' and stalma='" & Ctr_Ayuda1.xclave & "' "
        Set rabusca = VGCNx.Execute(nsql)
        If rabusca.RecordCount > 0 Then
          Label5 = Escadena(rabusca!adescri)
          txtcanti(1) = Round(rabusca!stskdis, 3)
        Else
          MsgBox "No existe articulo...!!", vbInformation, "AVISO"
          rabusca.Close
          Set rabusca = Nothing
          Exit Sub
        End If
        rabusca.Close
        cmdBotones(11).SetFocus
     End If
 End If
 Set rabusca = Nothing
 
End Sub

Private Sub TDBGrid1_Click()
   If rsdeta.RecordCount > 0 Then
      TDBGrid1.SetFocus
   End If
End Sub

Private Sub TDBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
  Dim nvalor As String
  If KeyCode = 46 Then
     If rsdeta.RecordCount <= 0 Then
        MBox2(1) = ""
        txtcanti(0) = "": txtcanti(1) = "": Label5 = ""
        Exit Sub
     End If
     nvalor = TDBGrid1.Columns(0).Text
     If rsdeta.RecordCount > 0 Then
        rsdeta.MoveFirst
        Do Until rsdeta.EOF
          If rsdeta.Fields(0) = nvalor Then
            rsdeta.Delete adAffectCurrent
            rsdeta.Update
            Exit Do
          End If
          rsdeta.MoveNext
        Loop
     End If
     ConfigGrid
     MBox2(1).SetFocus
     Exit Sub
          
'  ElseIf KeyCode = 13 Then
'    MBox2(1) = ""
'    txtcanti(0) = "": txtcanti(1) = "": Label5 = ""
'    MBox2(1) = TDBGrid1.Columns(1).Text
'    Label5 = TDBGrid1.Columns(2).Text
'    txtcanti(1) = Format(TDBGrid1.Columns(4).Text, "######0.000")
'    txtcanti(0) = Format(TDBGrid1.Columns(5).Text, "######0.000")
'    flag = 1
  End If
  
End Sub

Private Sub txtcanti_KeyPress(Index As Integer, KeyAscii As Integer)
  Dim posi As Integer
  If KeyAscii = 13 Then
    txtcanti(1) = Format(txtcanti(1), "##,###,##0.000")
    txtcanti(0) = Format(txtcanti(0), "##,###,##0.000")
    If Index = 0 Then
      SendKeys "{tab}"
    Else
      If rsdeta.RecordCount > 0 Then
        rsdeta.MoveLast
        posi = IIf(IsNull(rsdeta.Fields("item")), 0, rsdeta.Fields("item"))
      Else
        posi = 0
      End If
      txtcanti(0) = Format(txtcanti(0), "##,###,##0.0000")
      txtcanti(1) = Format(txtcanti(1), "##,###,##0.0000")
      If Val(txtcanti(1)) <= 0 Then
          MsgBox "Cantidad debe ser mayor a Cero..!!", vbInformation, "AVISO"
          Exit Sub
      End If
      If (dllgeneral.VerificaDatoExistente(VGCNx, "select * from stkart where stalma='" & Ctr_Ayuda1.xclave & "' and stcodigo='" & MBox2(1) & "' and round(stskdis,2)< " & txtcanti(1)) = 1) Then
        MsgBox " No existe saldo disponible...!!", vbInformation, "AVISO"
        Exit Sub
      ElseIf (dllgeneral.VerificaDatoExistente(VGCNx, "select * from stkart where stalma='" & Ctr_Ayuda1.xclave & "' and stcodigo='" & MBox2(1) & "'") = 0) Then
        MsgBox "No existe articulo en almacen...!!", vbInformation, "AVISO"
        Exit Sub
      End If

      rsdeta.AddNew
      rsdeta.Fields(0) = posi + 1
      rsdeta.Fields(1) = Escadena(MBox2(1))
      rsdeta.Fields(2) = Left(Escadena(Label5) & Space(40), 40)
      rsdeta.Fields(3) = ""
      rsdeta.Fields(4) = Format(txtcanti(1), "##,###,##0.000")
      rsdeta.Fields(5) = Format(txtcanti(0), "##,###,##0.000")
      rsdeta.Update
      MBox2(1) = ""
      txtcanti(0) = "": txtcanti(1) = "": Label5 = ""
      MBox2(1).SetFocus
    End If
  End If
End Sub
