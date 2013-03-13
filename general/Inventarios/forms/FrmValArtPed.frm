VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{4D137D9C-00A6-4458-9B46-7E95DB76D55B}#9.0#0"; "TextFer.ocx"
Begin VB.Form FrmValArtPed 
   Caption         =   "Valorización de Articulo Pendientes"
   ClientHeight    =   6825
   ClientLeft      =   2535
   ClientTop       =   2370
   ClientWidth     =   9720
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   6825
   ScaleWidth      =   9720
   Begin VB.Frame Frame4 
      Caption         =   "Totales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1095
      Left            =   120
      TabIndex        =   42
      Top             =   5640
      Width           =   6735
      Begin TextFer.TxFer TxSerie 
         Height          =   300
         Index           =   0
         Left            =   240
         TabIndex        =   43
         Top             =   600
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   529
         Object.CausesValidation=   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   3
         Text            =   ""
         ColorIlumina    =   -2147483624
         SaltarAlEnter   =   -1  'True
         Valor           =   ""
         ColorTextoAlEnfocar=   8454143
      End
      Begin TextFer.TxFer TxFer1 
         Height          =   300
         Left            =   1680
         TabIndex        =   44
         Top             =   600
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   529
         Object.CausesValidation=   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   3
         Text            =   ""
         ColorIlumina    =   -2147483624
         SaltarAlEnter   =   -1  'True
         Valor           =   ""
         ColorTextoAlEnfocar=   8454143
      End
      Begin TextFer.TxFer TxSerie 
         Height          =   300
         Index           =   1
         Left            =   3240
         TabIndex        =   45
         Top             =   600
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   529
         Object.CausesValidation=   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   3
         Text            =   ""
         ColorIlumina    =   -2147483624
         SaltarAlEnter   =   -1  'True
         Valor           =   ""
         ColorTextoAlEnfocar=   8454143
      End
      Begin TextFer.TxFer TxSerie 
         Height          =   300
         Index           =   2
         Left            =   5280
         TabIndex        =   46
         Top             =   600
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   529
         Object.CausesValidation=   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   3
         Text            =   ""
         ColorIlumina    =   -2147483624
         SaltarAlEnter   =   -1  'True
         Valor           =   ""
         ColorTextoAlEnfocar=   8454143
      End
      Begin VB.Label Label23 
         Caption         =   "Valor Imponible"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   50
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label23 
         Caption         =   "Valor IGV"
         Height          =   255
         Index           =   1
         Left            =   1800
         TabIndex        =   49
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label23 
         Caption         =   "Valor Inafecto"
         Height          =   255
         Index           =   2
         Left            =   3360
         TabIndex        =   48
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label23 
         Caption         =   "Valor Neto"
         Height          =   255
         Index           =   3
         Left            =   5400
         TabIndex        =   47
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "&Validar"
      Height          =   735
      Left            =   6960
      Picture         =   "FrmValArtPed.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   5880
      Width           =   840
   End
   Begin VB.CommandButton Command7 
      Caption         =   "&Salir"
      Height          =   735
      Left            =   7920
      Picture         =   "FrmValArtPed.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5880
      Width           =   840
   End
   Begin VB.Frame Frame2 
      Caption         =   "Buscar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   5520
      Left            =   225
      TabIndex        =   32
      Top             =   120
      Width           =   9435
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   2880
         Style           =   2  'Dropdown List
         TabIndex        =   40
         Top             =   210
         Width           =   2415
      End
      Begin TrueOleDBGrid70.TDBGrid TDBNota 
         Height          =   1680
         Left            =   165
         TabIndex        =   39
         Top             =   645
         Width           =   9120
         _ExtentX        =   16087
         _ExtentY        =   2963
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
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=2"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(5)=   "Column(1).Width=2725"
         Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2646"
         Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
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
      Begin VB.TextBox TxtBuscar 
         Height          =   315
         Left            =   600
         TabIndex        =   34
         Top             =   225
         Width           =   1455
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "FrmValArtPed.frx":0884
         Left            =   6000
         List            =   "FrmValArtPed.frx":088E
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   210
         Width           =   1455
      End
      Begin MSFlexGridLib.MSFlexGrid FG 
         Height          =   3090
         Left            =   135
         TabIndex        =   37
         Top             =   2400
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   5450
         _Version        =   393216
         AllowUserResizing=   1
      End
      Begin VB.Label Label11 
         Caption         =   "Almacen"
         Height          =   255
         Left            =   2160
         TabIndex        =   41
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label22 
         Caption         =   "Filtro"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label21 
         Caption         =   "Indice"
         Height          =   255
         Left            =   5400
         TabIndex        =   35
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Valorizado"
      ForeColor       =   &H80000007&
      Height          =   5490
      Left            =   225
      TabIndex        =   1
      Top             =   150
      Width           =   7500
      Begin VB.ComboBox Combo3 
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "FrmValArtPed.frx":08A7
         Left            =   5400
         List            =   "FrmValArtPed.frx":08B1
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   3165
         Width           =   1455
      End
      Begin VB.ComboBox Combo4 
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "FrmValArtPed.frx":08C5
         Left            =   2280
         List            =   "FrmValArtPed.frx":08D2
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   3480
         Width           =   1455
      End
      Begin VB.TextBox Text5 
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2280
         TabIndex        =   23
         Top             =   4200
         Width           =   1455
      End
      Begin VB.TextBox Text4 
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   5385
         TabIndex        =   21
         Top             =   3930
         Width           =   1455
      End
      Begin VB.TextBox Text3 
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2280
         TabIndex        =   20
         Top             =   3840
         Width           =   1455
      End
      Begin VB.TextBox Text2 
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   5400
         TabIndex        =   19
         Top             =   3570
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2280
         TabIndex        =   18
         Top             =   3120
         Width           =   1455
      End
      Begin VB.Label Label20 
         Caption         =   "Label20"
         Height          =   255
         Left            =   3000
         TabIndex        =   31
         Top             =   960
         Width           =   3855
      End
      Begin VB.Label Label19 
         Caption         =   "Label1"
         Height          =   255
         Left            =   3000
         TabIndex        =   28
         Top             =   1680
         Width           =   2055
      End
      Begin VB.Label Label18 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2280
         TabIndex        =   27
         Top             =   2760
         Width           =   4575
      End
      Begin VB.Label Label17 
         Caption         =   "Descripcion"
         Height          =   255
         Left            =   480
         TabIndex        =   26
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Label Label16 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2280
         TabIndex        =   25
         Top             =   2415
         Width           =   1455
      End
      Begin VB.Label Label15 
         Caption         =   "Codigo Art."
         Height          =   255
         Left            =   480
         TabIndex        =   24
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "Total"
         Height          =   255
         Index           =   3
         Left            =   480
         TabIndex        =   22
         Top             =   4200
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "Cantidad"
         Height          =   255
         Index           =   2
         Left            =   4080
         TabIndex        =   17
         Top             =   3960
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "Costo Unitario"
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   16
         Top             =   3840
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Documento"
         Height          =   255
         Left            =   480
         TabIndex        =   15
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Transaccion"
         Height          =   255
         Left            =   480
         TabIndex        =   14
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Proveedor"
         Height          =   255
         Left            =   480
         TabIndex        =   13
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "Doc Referencial"
         Height          =   255
         Left            =   480
         TabIndex        =   12
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "Serie"
         Height          =   255
         Left            =   480
         TabIndex        =   11
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Moneda"
         Height          =   255
         Left            =   4080
         TabIndex        =   10
         Top             =   3240
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "Factura"
         Height          =   255
         Left            =   480
         TabIndex        =   9
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "Conversion"
         Height          =   255
         Left            =   480
         TabIndex        =   8
         Top             =   3480
         Width           =   1575
      End
      Begin VB.Label Label9 
         Caption         =   "Tipo de Cambio"
         Height          =   255
         Index           =   0
         Left            =   4080
         TabIndex        =   7
         Top             =   3600
         Width           =   1335
      End
      Begin VB.Label Label10 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2280
         TabIndex        =   6
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label lbltransa 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2280
         TabIndex        =   5
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label12 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2280
         TabIndex        =   4
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label Label13 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2280
         TabIndex        =   3
         Top             =   1320
         Width           =   4215
      End
      Begin VB.Label Label14 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2280
         TabIndex        =   2
         Top             =   2040
         Width           =   1455
      End
   End
End
Attribute VB_Name = "FrmValArtPed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsql As String
Dim precio As Double
Dim CANTIDAD As Double
Dim tipcam As Double
Dim Rs As Recordset
Public rs1 As New ADODB.Recordset
Dim rsNota As ADODB.Recordset
Dim mRsql As String
Dim mRsql1 As String
Dim totdoc As Double
Dim sCodMon As String
Dim Fecha As Date   'Fecha del documento

  Dim i0 As Integer
  Dim xAlma As String
  Dim xDescri_alma As String
Dim rsSTKART As New ADODB.Recordset
Private Sub Combo2_Click()
    Call cargar_grid
End Sub


Private Sub Form_Load()
    Dim rsc As New ADODB.Recordset
    Set rsc = VGCNx.Execute("Select  TAALMA,TADESCRI  from  tabalm")
    If rsc.RecordCount > 0 Then
        Combo2.Clear
        rsc.MoveFirst
        Do Until rsc.EOF
            Combo2.AddItem rsc.Fields(0) & "-" & IIf(IsNull(rsc.Fields(1)), "****", rsc.Fields(1))
            rsc.MoveNext
        Loop
    End If
    rsc.Close
    Set rsc = Nothing
  central FrmValArtPed
  
    Label12 = ""
  Label13 = ""
  Label14 = ""
  Label19 = ""
  Text3 = ""
  Combo1.ListIndex = 0
  Combo3.ListIndex = 0
  Combo2.ListIndex = 0
  Combo4.ListIndex = 0
  
  i0 = InStr(Combo2.text, "-")
  xDescri_alma = Left(Combo2.text, i0 - 1)
  Frame1.Visible = False
   Call cargar_grid
  
End Sub

Private Sub Combo1_Click()
     FG.Col = Combo1.ListIndex
     FG.Sort = 5
End Sub

Private Sub Combo4_Click()
  If Combo4.ListIndex = 2 Then
     Text2.SetFocus
  End If
End Sub

Private Sub CmdAceptar_Click()
  Dim cant As String
  Dim Lote As String
  Dim Serie As String
  Dim uSql As String
  Dim rsql As String
  Dim codmon As String * 2
  
  '---------

    
  i0 = InStr(Combo2.text, "-")
  xDescri_alma = Left(Combo2.text, i0 - 1)
  '---------
  
  
  If Frame1.Visible Then
      If Not IsNumeric(Text3) Then
            MsgBox "Ingrese el Precio unitario !", vbOKOnly, "Error"
            Text3.SetFocus
            Exit Sub
     End If
     If Not IsNumeric(Text4) Then
            MsgBox "Ingrese la cantidad !", vbOKOnly, "Error"
            Text4.SetFocus
            Exit Sub
     End If
     If Combo3.ListIndex = 0 Then
        codmon = "01"
     Else
        codmon = "02"
     End If
     If sCodMon <> codmon Then
           If MsgBox("Desea Ud. cambiar el Tipo de moneda declarado inicialmente?", vbYesNo, "Aviso") = vbNo Then
                Exit Sub
           End If
     End If
     If Not IsNumeric(Text2) Then
           MsgBox "Ingrese el tipo de cambio !", vbOKOnly, "Error"
           Text2.SetFocus
           Exit Sub
     Else
           tipcam = Val(Text2)
     End If
     If Val(Text2) = 0 And codmon = "02" Then
           MsgBox "Ingrese el tipo de cambio !", vbOKOnly, "Error"
           Text2.SetFocus
           Exit Sub
     End If
     
     
     If codmon = "01" Then
        precio = Val(Text3.text) '* tipcam
     Else
        precio = Val(Text3.text) '* Val(Text2)
     End If
     CANTIDAD = Val(Text4.text)
     uSql = "Update MovAlmCab set CACODMON = '" & codmon & "', CATIPCAM = " & Val(Text2) & " where CANUMDOC='" & Trim(FG.TextMatrix(FG.Row, 3)) & "' and CAALMA = '" & xDescri_alma & "'    AND CATD='" & Trim(FG.TextMatrix(FG.Row, 2)) & "' "
     VGCNx.Execute uSql
     uSql = "Update MovAlmDet set DEPRECIO = " & precio & ",DETIPCAM = " & Val(Text2) & ",DECODMON = '" & codmon & "' where DENUMDOC='" & Trim(FG.TextMatrix(FG.Row, 3)) & "' and DECODIGO ='" & Trim(FG.TextMatrix(FG.Row, 0)) & "'and DEALMA = '" & xDescri_alma & "'  and  DETD='" & Trim(FG.TextMatrix(FG.Row, 2)) & "' "
     VGCNx.Execute uSql
     Call grabastk  'valoriza
     Call Totales
     Frame1.Visible = False
     Text4 = ""
     Text5 = ""
     limpiaGrid
     Frame2.Visible = True
     CmdAceptar.Caption = "&Validar"
  Else
     Text2 = "0"
     Text3 = "0"
     If FG.Rows = 1 Then Exit Sub
     Frame2.Visible = False
     Frame1.Visible = True
     CmdAceptar.Caption = "&Aceptar"
     rsql = "select  cacodmon,catipcam,cafecdoc from  MovAlmCab  where   CAALMA ='" & xDescri_alma & "'  and CATD='" & Trim(FG.TextMatrix(FG.Row, 2)) & "'  and CANUMDOC= '" & Trim(FG.TextMatrix(FG.Row, 3)) & "'" '    "'  n.DENUMDOC "
     
     Set Rs = VGCNx.Execute(rsql)
     If Not Rs.EOF Then
            If Rs("CACODMON") = "02" Then
                Combo3.ListIndex = 1
                sCodMon = "02"
            Else
                Combo3.ListIndex = 0
                sCodMon = "01"
            End If
            If Rs(1) <> 0 Then
                Text2 = Rs(1)
            End If
     Fecha = Rs(2)
     End If
     Rs.Close
     Label10 = FG.TextMatrix(FG.Row, 3)   'numdoc
     Lbltransa = FG.TextMatrix(FG.Row, 2)
     Label16 = FG.TextMatrix(FG.Row, 0)  ' çod
     Label18 = FG.TextMatrix(FG.Row, 1)
     Label13 = TDBNota.Columns(2)    ' proveedor
     Label12 = FG.TextMatrix(FG.Row, 5)     'RFTDOC
     Text1 = FG.TextMatrix(FG.Row, 6)
     If Label12 <> "" Then
        Label19 = tipref(Label12)
     End If
     If Lbltransa <> "" Then Label20 = transa(Lbltransa)
     Call cantidad_art(cant, Serie, Lote)
     Text4.Enabled = True
     Text4 = cant
     Text4.Enabled = False
     If Lote = "" Then
            Label14 = Serie
     Else
            Label14 = Lote
     End If
     Text3 = UltimoPrecio(Label16, sCodMon) 'precio Sugerido
     Text3.SetFocus
  End If
End Sub

Private Sub Command7_Click()
  If Frame1.Visible Then
        Frame1.Visible = False
        Frame2.Visible = True
        CmdAceptar.Caption = "&Validar"
        Text4 = ""
        Text5 = ""
  Else
        Unload Me
  End If
  
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
        rsSTKART.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
     Unload Me
End Sub

Private Sub TDBNota_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
  totdoc = 1
   Call cargar_grilla2
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
   SendKeys "{tab}"
 Else
   If ((Chr$(KeyAscii) > "9" Or Chr$(KeyAscii) < "0")) And Chr(KeyAscii) <> "." And KeyAscii <> 8 Then KeyAscii = 0
 End If
End Sub

Private Sub Text5_Change()
  If Text4 <> "" And IsNumeric(Text5) Then
             Text3 = Format(Val(Text5) / Val(Text4), "###0.0000")
  End If
End Sub
Private Sub Text3_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 And IsNumeric(Text3) Then
            CmdAceptar.SetFocus
            Exit Sub
  End If
  If ((Chr$(KeyAscii) > "9" Or Chr$(KeyAscii) < "0")) And Chr(KeyAscii) <> "." And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If IsNumeric(Text4) And KeyAscii = 13 And IsNumeric(Text3) Then
        If Not IsNumeric(Text3) Then Exit Sub
Text5 = Val(Text3) * Val(Text4)
ElseIf KeyAscii = 13 And IsNumeric(Text5) And IsNumeric(Text4) Then
        Text3 = Format(Val(Text5) / Val(Text4), "##0.0000")
Else
        If Chr$(KeyAscii) = "." Then Exit Sub
        If ((Chr$(KeyAscii) > "9" Or Chr$(KeyAscii) < "0")) And KeyAscii <> 8 Then KeyAscii = 0
End If
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And IsNumeric(Text5) And IsNumeric(Text4) Then
      Text3 = Val(Text5) / Val(Text4)
      Text3 = Format(Text3, "##0.0000")
      CmdAceptar.SetFocus
Else
      If Chr$(KeyAscii) = "." Then Exit Sub
      If ((Chr$(KeyAscii) > "9" Or Chr$(KeyAscii) < "0")) And KeyAscii <> 8 Then KeyAscii = 0
End If
End Sub

Public Sub grabastk()
   Dim criterio As String
   Dim cadena As String
   Dim auxdisp As Double
   Dim AUXPRECIO As Double
   cadena = Label16
   criterio = " STCODIGO ='" & cadena & "' and  STALMA ='" & xDescri_alma & "'"
   rsSTKART.Filter = criterio
        
   If Combo3.ListIndex = 0 Then
       AUXPRECIO = precio
   Else
       If Val(Text2) <> 0 Then
          AUXPRECIO = precio * Val(Text2)
       End If
   End If
   
   If Not rsSTKART.EOF Then

     auxdisp = rsSTKART("STSKDIS")
     If rsSTKART("STKPREPRO") <> 0 And (CANTIDAD + auxdisp) <> 0 Then   'no se registrado algun precio
         rsSTKART("STKPREPRO") = (AUXPRECIO * CANTIDAD + auxdisp * rsSTKART("STKPREPRO")) / (CANTIDAD + auxdisp)
        If IsNull(rsSTKART("stkultfechacompra")) Or (rsSTKART("stkultfechacompra") <= Fecha) Then rsSTKART("stkultfechacompra") = Fecha
     Else

        rsSTKART("STKPREPRO") = AUXPRECIO

     End If
     If IsNull(rsSTKART("STKFECULT")) Or (rsSTKART("STKFECULT") <= Fecha) Then
        rsSTKART("STKFECULT") = Fecha
        rsSTKART("STKPREULT") = AUXPRECIO '*RMM***********  'Precio
     End If
   End If
   rsSTKART.Update
  
End Sub

Private Sub cantidad_art(pcantidad As String, pserie As String, plote As String)
 Dim Rs As Recordset
 Dim rsql As String
 
 rsql = "select decantid,delote,deserie from MovAlmdet where DENUMDOC='" & Trim(FG.TextMatrix(FG.Row, 3)) & "' and DECODIGO ='" & Trim(FG.TextMatrix(FG.Row, 0)) & "' and DEALMA = '" & xDescri_alma & "' AND  DETD='" & Trim(FG.TextMatrix(FG.Row, 2)) & "'"
 
 Set Rs = VGCNx.Execute(rsql)
  If Rs.EOF Then
            pcantidad = "0"
            pserie = ""
            plote = ""
  Else
            pcantidad = Str(Rs(0))
            pserie = IIf(Not IsNull(Rs(1)), Rs(1), "")
            pserie = IIf(Not IsNull(Rs(2)), Rs(2), "")
  End If
End Sub

Function tipref(text As Label) As String
 
 Dim Rs As Recordset
 Dim rsql As String
 rsql = "select  TDO_DESCRI FROM TIPO_DOCU  where TDO_TIPDOC= '" & text & "'" '
 Set Rs = VGCNx.Execute(rsql)
 tipref = IIf(Not Rs.EOF, Rs(0), "")
 Rs.Close
End Function

Function transa(text As Label) As String
 Dim Rs As Recordset
 Dim rsql As String
 Dim dato As String
  dato = "I"
  rsql = "select  TT_DESCRI FROM TabTransa where TT_CODMOV= '" & text & "' and TT_TIPMOV ='" & dato & "'" '
  Set Rs = VGCNx.Execute(rsql)
  transa = IIf(Not Rs.EOF, Rs(0), "")
  Rs.Close
End Function

Private Sub limpiaGrid()
Dim i As Integer
 If FG.Rows = 1 Then Exit Sub
 i = FG.RowSel
 If FG.Rows > 2 Then
        FG.RemoveItem i
 Else
        FG.Clear
        FG.Rows = 1
        FG.FormatString = "Cod. Articulo.|Descripcion| Tr| Num.Doc."
        FG.Row = 0
        FG.ColWidth(0) = 950
        FG.ColWidth(1) = 3700
        FG.ColWidth(2) = 450
        FG.ColWidth(3) = 1300
        FG.ColWidth(4) = 2
        FG.ColWidth(5) = 2
  End If
End Sub



Private Sub Txtbuscar_Change()
Dim i As Integer
Dim n As Integer
n = Combo1.ListIndex
If TxtBuscar <> "" Then
      For i = 1 To FG.Rows - 1
          If UCase(Left(FG.TextMatrix(i, n), Len(TxtBuscar))) = UCase(Trim(TxtBuscar)) Then
             Exit For
          End If
      Next i
      If i >= FG.Rows Then
            FG.HighLight = flexHighlightNever
      Else
            FG.HighLight = flexHighlightAlways
            FG.TopRow = i
            FG.Row = i
            FG.Col = 0
            FG.ColSel = FG.Cols - 1
      End If
End If
End Sub
Public Sub cargar_grid()

   i0 = InStr(Combo2.text, "-")
   xDescri_alma = Left(Combo2.text, i0 - 1)
       '****************************************************RMM 07/07/2001
  Set rsSTKART = New ADODB.Recordset

  rsSTKART.Open "Select * from STKART WHERE STALMA='" & xDescri_alma & "'", VGCNx, adOpenDynamic, adLockOptimistic

  Dim sqlcad As String
  
  sqlcad = "select  N.DETD as TD,n.DENUMDOC as 'Num. Doc.', p.clienterazonsocial as Proveedor ,"
  sqlcad = sqlcad & "m.CARFTDOC as 'Doc. Ref.', m.CARFNDOC as 'Nro. Refe.',cafecdoc as 'Fecha Doc' from MovAlmCab m "
  sqlcad = sqlcad & " inner join MovAlmDet n on n.DEALMA = m.CAALMA and n.DENUMDOC = m.CANUMDOC  and n.DETD= m.CATD "
  sqlcad = sqlcad & " inner join MaeArt x on n.DECODIGO=x.ACODIGO   "
  sqlcad = sqlcad & " left join cp_proveedor p on m.cacodpro=p.clientecodigo  Where  m.CAALMA ='" & xDescri_alma & "' AND (CATD='NI' OR CATD='NC' ) and  n.DEPRECIO = 0  "
  sqlcad = sqlcad & " and  m.casitgui<>'A'  group by N.DETD,n.DENUMDOC, p.clienterazonsocial ,m.CARFTDOC, m.CARFNDOC,m.cafecdoc order by 1,2"
  
  Set rsNota = New ADODB.Recordset
  Set rsNota = VGCNx.Execute(sqlcad)
  
    If rsNota.RecordCount = 0 Then
        MsgBox "No hay Notas de Ingreso/Salida", vbInformation, Caption
        rsNota.Close
        Set TDBNota.DataSource = Nothing
        FG.Clear
        'Unload Me
        Exit Sub
    End If

  
  Set TDBNota.DataSource = rsNota
  
  TDBNota.Columns(0).Width = 600
  TDBNota.Columns(1).Width = 1500
  TDBNota.Columns(2).Width = 3800
  TDBNota.Columns(3).Width = 600
  TDBNota.Columns(4).Width = 1500
 
  mRsql1 = "select n.STCODIGO FROM  StkArt n where n.STALMA = '" & xDescri_alma & "'"
  Set rs1 = VGCNx.Execute(mRsql1)

  
End Sub

Public Sub cargar_grilla2()

          mRsql = "select  n.DECODIGO, ADESCRI, N.DETD,n.DENUMDOC, m.CANOMPRO, m.CARFTDOC, m.CARFNDOC from MovAlmCab m, MovAlmDet n ,MaeArt  Where  m.CAALMA ='" & xDescri_alma & _
                             "' AND n.DEALMA = m.CAALMA and (CATD='NI' OR CATD='NC' )   and  n.DEPRECIO = 0   and ACODIGO  = n.DECODIGO      And   n.DENUMDOC = m.CANUMDOC  and n.DETD= m.CATD and m.CASITGUI<>'A'  AND "
          mRsql = mRsql & "n.DETD='" & TDBNota.Columns(0).Value & "' and n.DENUMDOC='" & TDBNota.Columns(1).Value & "' ORDER BY m.CANUMDOC"
        
          Set Rs = VGCNx.Execute(mRsql)
          If Rs.RecordCount = 0 Then
                     MsgBox "No hay Artículos por Valorizar que esten Pendientes", vbExclamation, mensaje1
                     FG.Clear
          Else
                    
                    Call limpiar_grilla2
                    FG.Rows = 1
                    Rs.MoveFirst
                    FG.Visible = False
                    While Not Rs.EOF
                            FG.AddItem (Rs(0) & vbTab & Trim(Rs(1)) & vbTab & Rs(2) & vbTab & Rs(3) & vbTab & Rs(4) & vbTab & Rs(5) & vbTab & Rs(6))
                            Rs.MoveNext
                    Wend
                    Rs.Close
                    FG.Visible = True
          End If
End Sub

Public Sub limpiar_grilla2()

    FG.Clear
    FG.Cols = 7
    'FG.FormatString = "Codigo Art.|Descripcion| TD |Num.Doc| |"
    FG.Row = 0
    FG.ColWidth(0) = 1400
    FG.ColWidth(1) = 5100
    FG.ColWidth(2) = 500
    FG.ColWidth(3) = 1000
    FG.ColWidth(4) = 1000
    FG.ColWidth(5) = 800
    FG.ColWidth(6) = 1000
    
    'FG.FormatString = "Codigo Art.|Descripcion| TD |Num.Doc| |"
    FG.ColAlignment(0) = 1
    FG.ColAlignment(1) = 1
    
    FG.Row = 0
    FG.Col = 0
    Dim cabecera(1, 6)
    Dim i As Integer
    i = 0
    cabecera(1, 0) = "Codigo"
    cabecera(1, 1) = "Descripcion"
    cabecera(1, 2) = "TD"
    cabecera(1, 3) = "Num. Doc" '--"Nro. Documento"
    cabecera(1, 4) = "Proveedor" '---"Proveedor"
    cabecera(1, 5) = "Doc. Ref" '--"Doc.]REf"
    cabecera(1, 6) = "Num. Ref." '- -"Num ref"
    
    
    For i = 0 To FG.Cols - 1
        FG.Col = i
        FG.text = cabecera(1, i)
    Next i

End Sub

Private Sub Totales()
rsql = " select decantid,deprecio,decodigo from movalmdet inner join movalmcab on dealma+detd+denumdoc=caalma+catd+canumdoc where dealma='" & xDescri_alma & "' and  "
rsql = rsql & " DETD='" & TDBNota.Columns(0).Value & "' and DENUMDOC='" & TDBNota.Columns(1).Value & "' and CASITGUI<>'A'"
Set rs1 = VGCNx.Execute(rsql)
TxSerie(2).text = 0
Do Until rs1.EOF()
   TxSerie(2).text = TxSerie(2).text + rs1!DECANTID * ESNULO(rs1!DEPRECIO, 0)
   rs1.MoveNext
Loop
Call cargar_grilla2
End Sub
