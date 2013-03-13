VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Begin VB.Form FrmGuiaSal 
   Caption         =   "Form2"
   ClientHeight    =   7635
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   12480
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   7635
   ScaleWidth      =   12480
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FrameRipley 
      BackColor       =   &H0080FFFF&
      Caption         =   "Aceptar"
      Height          =   2295
      Left            =   3480
      TabIndex        =   86
      Top             =   2760
      Visible         =   0   'False
      Width           =   8415
      Begin VB.CommandButton CmdSair 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Salir"
         Height          =   495
         Left            =   4200
         MaskColor       =   &H00C0C0FF&
         TabIndex        =   94
         Top             =   1680
         Width           =   1215
      End
      Begin VB.CommandButton CmdAceptar 
         Caption         =   "Aceptar"
         Height          =   495
         Left            =   1800
         TabIndex        =   93
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox TextDNI 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   960
         TabIndex        =   89
         Top             =   240
         Width           =   1416
      End
      Begin VB.TextBox TextRazon 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         MaxLength       =   40
         TabIndex        =   88
         Top             =   720
         Width           =   4410
      End
      Begin VB.TextBox TextDir 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   960
         MaxLength       =   50
         TabIndex        =   87
         Top             =   1200
         Width           =   7290
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "DNI-RUC"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   92
         Top             =   240
         Width           =   645
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Nombre-Raz.Soc."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   91
         Top             =   720
         Width           =   1380
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Direccion"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   90
         Top             =   1200
         Width           =   645
      End
   End
   Begin VB.TextBox TxtTransp 
      Height          =   330
      Left            =   1695
      TabIndex        =   85
      Top             =   5415
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Frame FrameComentario 
      Caption         =   "Comentarios"
      Height          =   2100
      Left            =   2520
      TabIndex        =   81
      Top             =   3750
      Visible         =   0   'False
      Width           =   8316
      Begin VB.TextBox TxComentario 
         Height          =   1695
         Left            =   600
         MultiLine       =   -1  'True
         TabIndex        =   84
         Top             =   240
         Width           =   5655
      End
      Begin VB.CommandButton CmdComGrabar 
         Caption         =   "Grabar"
         Height          =   495
         Left            =   6600
         TabIndex        =   83
         Top             =   600
         Width           =   1335
      End
      Begin VB.CommandButton CmdComCan 
         Caption         =   "Cancelar"
         Height          =   495
         Left            =   6600
         TabIndex        =   82
         Top             =   1320
         Width           =   1335
      End
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Height          =   1215
      Index           =   0
      Left            =   2670
      TabIndex        =   75
      Top             =   6090
      Width           =   7515
      Begin VB.CommandButton Command1 
         Caption         =   "&Adicionar"
         Height          =   1005
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   80
         Top             =   90
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Modificar"
         Height          =   1005
         Left            =   1620
         Style           =   1  'Graphical
         TabIndex        =   79
         Top             =   90
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Eliminar"
         Height          =   1005
         Left            =   3180
         Style           =   1  'Graphical
         TabIndex        =   78
         Top             =   90
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.CommandButton CmdGrabarDet 
         Caption         =   "&Grabar"
         Height          =   1005
         Left            =   4755
         Style           =   1  'Graphical
         TabIndex        =   77
         Top             =   90
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.CommandButton CmdSalir 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         Height          =   1005
         Left            =   6300
         Style           =   1  'Graphical
         TabIndex        =   76
         Top             =   90
         Width           =   1155
      End
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Height          =   525
      Index           =   1
      Left            =   360
      TabIndex        =   70
      Top             =   5490
      Width           =   11805
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   1410
         TabIndex        =   72
         Top             =   120
         Width           =   1395
      End
      Begin VB.TextBox Text9 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   10140
         TabIndex        =   71
         Top             =   90
         Width           =   1425
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Total  Items :"
         Height          =   195
         Index           =   1
         Left            =   210
         TabIndex        =   74
         Top             =   150
         Width           =   1125
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Total  Cantidad :"
         Height          =   195
         Index           =   0
         Left            =   8715
         TabIndex        =   73
         Top             =   150
         Width           =   1365
      End
   End
   Begin VB.Frame FrmPen 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Caption         =   "Pendientes"
      Height          =   1875
      Left            =   2415
      TabIndex        =   16
      Top             =   4590
      Visible         =   0   'False
      Width           =   8400
      Begin VB.CommandButton CmdCan 
         Caption         =   "Cancelar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6975
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   45
         Width           =   1275
      End
      Begin TrueOleDBGrid70.TDBGrid TDBGrid 
         Height          =   1380
         Left            =   135
         TabIndex        =   17
         Top             =   405
         Width           =   8160
         _ExtentX        =   14393
         _ExtentY        =   2434
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
         Splits(0).RecordSelectorWidth=   688
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   14215660
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
         DeadAreaBackColor=   14215660
         RowDividerColor =   14215660
         RowSubDividerColor=   14215660
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
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Guias Pendientes"
         Height          =   195
         Left            =   150
         TabIndex        =   19
         Top             =   120
         Width           =   1455
      End
   End
   Begin VB.Frame FrmValida 
      BackColor       =   &H00C9955A&
      BorderStyle     =   0  'None
      Caption         =   "Pendientes"
      Height          =   2580
      Left            =   2430
      TabIndex        =   0
      Top             =   2325
      Visible         =   0   'False
      Width           =   8940
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   300
         Left            =   8280
         Top             =   1530
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H80000009&
         Caption         =   "Cancelar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7470
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   225
         Width           =   1275
      End
      Begin TrueOleDBGrid70.TDBGrid GridP 
         Height          =   1650
         Left            =   180
         TabIndex        =   2
         Top             =   585
         Width           =   8610
         _ExtentX        =   15187
         _ExtentY        =   2910
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
         Splits(0).RecordSelectorWidth=   688
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   14215660
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
         DeadAreaBackColor=   14215660
         RowDividerColor =   14215660
         RowSubDividerColor=   14215660
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
      Begin VB.Shape Shape1 
         BorderColor     =   &H80000009&
         BorderWidth     =   3
         Height          =   2490
         Left            =   45
         Top             =   45
         Width           =   8835
      End
      Begin VB.Label Lblmensaje 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "El stock de los siguientes productos no cubren la cantidad pedida"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   195
         TabIndex        =   15
         Top             =   255
         Width           =   6330
      End
      Begin VB.Label LblPro 
         Caption         =   "Label15"
         Height          =   330
         Index           =   0
         Left            =   450
         TabIndex        =   14
         Top             =   2745
         Width           =   1365
      End
      Begin VB.Label LblFal 
         Caption         =   "Label15"
         Height          =   330
         Index           =   0
         Left            =   2160
         TabIndex        =   13
         Top             =   2790
         Width           =   2310
      End
      Begin VB.Label LblPro 
         Caption         =   "Label15"
         Height          =   330
         Index           =   1
         Left            =   450
         TabIndex        =   12
         Top             =   3105
         Width           =   1365
      End
      Begin VB.Label LblPro 
         Caption         =   "Label15"
         Height          =   330
         Index           =   2
         Left            =   405
         TabIndex        =   11
         Top             =   3465
         Width           =   1365
      End
      Begin VB.Label LblPro 
         Caption         =   "Label15"
         Height          =   330
         Index           =   3
         Left            =   405
         TabIndex        =   10
         Top             =   3825
         Width           =   1365
      End
      Begin VB.Label LblPro 
         Caption         =   "Label15"
         Height          =   330
         Index           =   4
         Left            =   405
         TabIndex        =   9
         Top             =   4230
         Width           =   1365
      End
      Begin VB.Label LblPro 
         Caption         =   "Label15"
         Height          =   330
         Index           =   5
         Left            =   405
         TabIndex        =   8
         Top             =   4680
         Width           =   1365
      End
      Begin VB.Label LblFal 
         Caption         =   "Label15"
         Height          =   330
         Index           =   1
         Left            =   2115
         TabIndex        =   7
         Top             =   3195
         Width           =   2310
      End
      Begin VB.Label LblFal 
         Caption         =   "Label15"
         Height          =   330
         Index           =   2
         Left            =   2070
         TabIndex        =   6
         Top             =   3555
         Width           =   2310
      End
      Begin VB.Label LblFal 
         Caption         =   "Label15"
         Height          =   330
         Index           =   3
         Left            =   2025
         TabIndex        =   5
         Top             =   3870
         Width           =   2310
      End
      Begin VB.Label LblFal 
         Caption         =   "Label15"
         Height          =   330
         Index           =   4
         Left            =   2025
         TabIndex        =   4
         Top             =   4230
         Width           =   2310
      End
      Begin VB.Label LblFal 
         Caption         =   "Label15"
         Height          =   330
         Index           =   5
         Left            =   2025
         TabIndex        =   3
         Top             =   4590
         Width           =   2310
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   1740
      Left            =   360
      TabIndex        =   20
      Top             =   3750
      Width           =   11820
      _ExtentX        =   20849
      _ExtentY        =   3069
      _Version        =   393216
      Rows            =   1
      Cols            =   7
      FormatString    =   $"FrmGuiaSal.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   3510
      Left            =   360
      TabIndex        =   21
      Top             =   120
      Width           =   11805
      Begin VB.TextBox Text13 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9045
         TabIndex        =   45
         Top             =   180
         Visible         =   0   'False
         Width           =   2025
      End
      Begin VB.TextBox TxtAlmacen 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   10665
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   43
         Top             =   1290
         Width           =   495
      End
      Begin VB.TextBox Text8 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1470
         Locked          =   -1  'True
         TabIndex        =   42
         Top             =   1950
         Width           =   1335
      End
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1470
         TabIndex        =   41
         Top             =   1290
         Width           =   4590
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7785
         TabIndex        =   40
         Top             =   930
         Width           =   3915
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1470
         TabIndex        =   39
         Top             =   930
         Width           =   1305
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7635
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   38
         Top             =   570
         Width           =   1335
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1470
         MaxLength       =   2
         TabIndex        =   37
         Top             =   570
         Width           =   645
      End
      Begin VB.TextBox TxTransa 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1470
         MaxLength       =   2
         TabIndex        =   36
         Top             =   210
         Width           =   645
      End
      Begin VB.CommandButton CmdGrabarCab 
         Caption         =   ">>"
         Height          =   255
         Left            =   11250
         TabIndex        =   35
         Top             =   1320
         Width           =   435
      End
      Begin VB.ComboBox CmbSerie 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5610
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   570
         Width           =   1035
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5055
         Locked          =   -1  'True
         TabIndex        =   33
         Top             =   1950
         Width           =   1416
      End
      Begin VB.CheckBox ChkTalla 
         Alignment       =   1  'Right Justify
         Caption         =   "Salidas por Talla"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9300
         TabIndex        =   32
         Top             =   1800
         Width           =   1680
      End
      Begin VB.TextBox tx_codmaq 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1470
         TabIndex        =   31
         Top             =   2685
         Width           =   1416
      End
      Begin VB.TextBox tx_ordfab 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7770
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   1710
         Width           =   1416
      End
      Begin VB.CommandButton Bdire 
         Caption         =   "..."
         Height          =   285
         Left            =   6210
         TabIndex        =   29
         Top             =   1275
         Width           =   375
      End
      Begin VB.TextBox Txtnrodoc 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7785
         Locked          =   -1  'True
         MaxLength       =   11
         TabIndex        =   28
         Top             =   1290
         Width           =   1416
      End
      Begin VB.TextBox Texttipdoc 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3060
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   27
         Top             =   1950
         Width           =   495
      End
      Begin VB.Frame Frame2 
         Caption         =   "Direcciones"
         Height          =   1740
         Left            =   6750
         TabIndex        =   23
         Top             =   3330
         Visible         =   0   'False
         Width           =   4995
         Begin VB.CommandButton cCerrar 
            Caption         =   "Cerrar"
            Height          =   345
            Left            =   5550
            TabIndex        =   26
            Top             =   2160
            Width           =   1365
         End
         Begin VB.CommandButton cAcepta 
            Caption         =   "&Acepta"
            Height          =   345
            Left            =   4260
            TabIndex        =   25
            Top             =   2160
            Width           =   1155
         End
         Begin MSDataGridLib.DataGrid dbGrid1 
            Height          =   1380
            Left            =   135
            TabIndex        =   24
            Top             =   225
            Width           =   4560
            _ExtentX        =   8043
            _ExtentY        =   2434
            _Version        =   393216
            HeadLines       =   1
            RowHeight       =   15
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   10250
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   10250
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               MarqueeStyle    =   3
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
      End
      Begin VB.TextBox TxtCon 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1470
         MaxLength       =   50
         TabIndex        =   22
         Top             =   1620
         Width           =   4590
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   315
         Left            =   7605
         TabIndex        =   44
         Top             =   180
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   39321601
         CurrentDate     =   36437
         MaxDate         =   401768
         MinDate         =   36161
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyuVendedor 
         Height          =   315
         Left            =   1470
         TabIndex        =   46
         Top             =   2295
         Width           =   3795
         _ExtentX        =   6694
         _ExtentY        =   556
         XcodMaxLongitud =   3
         xcodwith        =   200
         NomTabla        =   "vt_vendedor"
         TituloAyuda     =   "Ayuda de Vendedores"
         ListaCampos     =   "vendedorcodigo(1),vendedornombres(1)"
         XcodCampo       =   "vendedorcodigo"
         XListCampo      =   "vendedornombres"
         ListaCamposDescrip=   "Codigo,Descripcion"
         ListaCamposText =   "vendedorcodigo,vendedornombres"
         Requerido       =   0   'False
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyuTransporte 
         Height          =   315
         Left            =   6630
         TabIndex        =   47
         Top             =   2310
         Width           =   4770
         _ExtentX        =   8414
         _ExtentY        =   556
         XcodMaxLongitud =   11
         xcodwith        =   800
         NomTabla        =   "al_transporte"
         TituloAyuda     =   "Ayuda de Transporte"
         ListaCampos     =   "TRACODIGO(1),TRANOMBRE(1)"
         XcodCampo       =   "TRACODIGO"
         XListCampo      =   "TRANOMBRE"
         ListaCamposDescrip=   "Codigo,Descripcion"
         ListaCamposText =   "TRACODIGO,TRANOMBRE"
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyudaEmpresa 
         Height          =   450
         Left            =   1485
         TabIndex        =   48
         Top             =   900
         Visible         =   0   'False
         Width           =   5085
         _ExtentX        =   8969
         _ExtentY        =   794
         XcodMaxLongitud =   11
         xcodwith        =   1000
         NomTabla        =   "cp_proveedor"
         TituloAyuda     =   "Ayuda de Proveedores"
         ListaCampos     =   "clientecodigo(1),clienteruc(1),clienterazonsocial(1),clientedireccion(1)"
         XcodCampo       =   "clientecodigo"
         XListCampo      =   "clienterazonsocial"
         ListaCamposDescrip=   "Codigo,Descripcion"
         ListaCamposText =   "clientecodigo,clienteruc,clienterazonsocial,clientedireccion"
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Dirección"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   180
         TabIndex        =   69
         Top             =   1320
         Width           =   645
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Razon Social :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6705
         TabIndex        =   68
         Top             =   945
         Width           =   1005
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Num. Doc :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6795
         TabIndex        =   67
         Top             =   630
         Width           =   795
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Transacción :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   180
         TabIndex        =   66
         Top             =   270
         Width           =   960
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Doc. :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6615
         TabIndex        =   65
         Top             =   240
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Doc. Referencia :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   180
         TabIndex        =   64
         Top             =   630
         Width           =   1260
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Cod. Cliente :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   180
         TabIndex        =   63
         Top             =   990
         Width           =   990
      End
      Begin VB.Label Label19 
         Caption         =   "N°"
         Height          =   255
         Left            =   7770
         TabIndex        =   62
         Top             =   1020
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Transportista :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5340
         TabIndex        =   61
         Top             =   2340
         Width           =   945
      End
      Begin VB.Label lblSerie 
         Caption         =   "Serie :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5040
         TabIndex        =   60
         Top             =   630
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "Vendedor :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   180
         TabIndex        =   59
         Top             =   2310
         Width           =   795
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "Centro Costo :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   195
         Left            =   3750
         TabIndex        =   58
         Top             =   2010
         Width           =   1065
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Orden Compra :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   180
         TabIndex        =   57
         Top             =   2010
         Width           =   1155
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "Orden Fabricación :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   195
         Left            =   6240
         TabIndex        =   56
         Top             =   1770
         Visible         =   0   'False
         Width           =   1410
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "Equip./Maqui :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   195
         Left            =   180
         TabIndex        =   55
         Top             =   2700
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.Label LblCC 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8610
         TabIndex        =   54
         Top             =   2220
         Width           =   2910
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         BackColor       =   &H00C9955A&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   285
         Left            =   2190
         TabIndex        =   53
         Top             =   210
         Width           =   4245
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         BackColor       =   &H00C9955A&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   285
         Left            =   2190
         TabIndex        =   52
         Top             =   570
         Width           =   2535
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Nro Pedido :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6840
         TabIndex        =   51
         Top             =   1320
         Width           =   885
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Almacen Destino :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   9315
         TabIndex        =   50
         Top             =   1320
         Width           =   1230
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Contacto :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   180
         TabIndex        =   49
         Top             =   1620
         Width           =   765
      End
   End
End
Attribute VB_Name = "FrmGuiaSal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Este formulario es utilizado como Guia de Remision o Devolucion cualquier
'cambio se tiene que validar en ambos formulario
'VGGuiaSal es la variable global que indica si es guia o devolucion

Option Explicit
''Dim db As Database
Dim TCamb As Double
Dim hubo_error As Boolean
Dim direccion  As String
Dim xserie As String * 1
Dim VGDllGeneral As New dllgeneral.dll_general
Dim rg As New ADODB.Recordset

'Dim DbAux As Database
Dim numsal As Double       'Numero consecutivo de guia de remision
Dim Unid As String
Dim wok As Integer
Dim nument As Long       'Numero consecutivo de nota de ingreso
Dim precioprom As Double
Dim CANTIDAD As Double
Dim canttemp As Double
Dim Campo As String * 2  'Indica el tipo de transaccion
Dim contador As Long     'Indica el item del flex
Dim auxdisp As Long
Dim cantidadDEV As Double
Dim numserie As String    'Numero de guia remision
Dim tipofactura As String  ' tipo de docuemnto
Dim nroguia As String  '  numero de factura/boleta
Dim tipo, Codigo2 As String
Dim TT_CONTADOR As Integer
Dim salir As Boolean      'Numero de la serie de guia de remision
Dim Serie As String
Dim AlmacenRF As String
Dim cTransa As String
Dim ruc As String
Dim EstadoDevolucion As String
Dim WithEvents Conex As ADODB.Connection
Attribute Conex.VB_VarHelpID = -1
Dim Completo As Boolean
Dim flagserie, flaglote As String * 1
Dim dato_invalido As Boolean
Dim serie_lote As String
Dim rs As ADODB.Recordset  'agregado
Dim Rs2 As ADODB.Recordset  'agregado
Dim Cliente As Boolean, Requerimiento As Boolean
Dim analitico As Integer
Dim t As Integer

Private Sub pro_xserie()
  If flagserie = "S" Then
        xserie = "S"
        Exit Sub
  End If
  If flaglote = "S" Then
        xserie = "N"
        Exit Sub
  End If
  xserie = "X"
End Sub
Function coduso(dato As String) As String
   Dim RSQL As String
   Dim rs As New ADODB.Recordset
   RSQL = "select UM_ABREV from TabUniMed where UM_NOMBRE ='" & dato & "'"
   'Set RS = VGBaseDatos.OpenRecordset(RSQL, dbOpenSnapshot)
   Set rs = VGCNx.Execute(RSQL)
   If rs.RecordCount = 0 Then
    coduso = ""
   Else
    coduso = rs(0)
   End If
   rs.Close
   Set rs = Nothing
End Function

Function Nombre_Unidad(dato As String) As String
   Dim RSQL As String
   Dim rs As New ADODB.Recordset
   RSQL = "select UM_NOMBRE from TabUniMed where UM_ABREV ='" & dato & "'" '
   'Set RS = VGBaseDatos.OpenRecordset(RSQL, dbOpenSnapshot)
   Set rs = VGCNx.Execute(RSQL)
   If rs.RecordCount = 0 Then
     Nombre_Unidad = ""
   Else
     Nombre_Unidad = rs(0)
   End If
   rs.Close
   Set rs = Nothing
End Function
Private Sub limpia()
     
End Sub

Private Sub Bdire_Click()
On Error Resume Next

Set rg = Nothing
Set DbGrid1.DataSource = Nothing

Set rg = VGCNx.Execute("select cliedirnumero as Nro,cliedirdireccion as Direccion from vt_clientedireccion where clientecodigo='" & Text5 & "'")
If rg.RecordCount > 0 Then
    Frame2.Visible = True
    Set DbGrid1.DataSource = rg
    DbGrid1.Refresh
Else
   Frame2.Visible = False
End If
   
End Sub

Private Sub cAcepta_Click()
   Text7 = IIf(IsNull(DbGrid1.Columns(1).text), "", DbGrid1.Columns(1).text)
   Frame2.Visible = False
   Text7.SetFocus
End Sub

Private Sub cCerrar_Click()
   Frame2.Visible = False
   Text7.SetFocus
End Sub

Private Sub CmbSerie_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
   SendKeys "{tab}"
   KeyAscii = 0
 End If
End Sub

Private Sub cmdAceptar_Click()
If Len(RTrim(TextDNI)) = 0 Or Len(RTrim(TextRazon)) = 0 Or Len(RTrim(TextDir)) = 0 Then
   If MsgBox(" DNI / Nombre / Direccion en Blanco, desea continuar  ", vbYesNo + vbQuestion, "Mensaje ") = vbYes Then
      wok = 1
      FrameRipley.Visible = False
 Else
       wok = 1
      FrameRipley.Visible = False
   End If
  Else
      wok = 1
      FrameRipley.Visible = False
 End If
End Sub

Private Sub CmdCan_Click()
FrmPen.Visible = False
End Sub

Private Sub CmdComCan_Click()
'CANCELA EL COMENTARIO
Dim rpta As Integer
FrameComentario.Visible = False
crtlvisible (True)
rpta = MsgBox("¿Desea Imprimir?", vbYesNo + vbQuestion, "Aviso")
If rpta = vbYes Then
  '  imprimir
  imprimirguias
End If
End Sub

Private Sub CmdComGrabar_Click()
'GRABA EL COMENTARIO DE LA GUIA
Dim RSQL As String
Dim rpta As String
 
RSQL = "Update MovAlmCab set CAGLOSA = '" & TxComentario & "' "
RSQL = RSQL & "Where  CAALMA = '" & TxtAlmacen.text & "'AND  CATD= '" & tipo & "' AND CANUMDOC = '" & numserie & "'" '
VGCNx.Execute RSQL
FrameComentario.Visible = False
crtlvisible (True)
rpta = MsgBox("¿Desea Imprimir?", vbYesNo + vbQuestion, "Aviso")
If rpta = vbYes Then
    ' imprimir
    imprimirguias
End If
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If
End Sub

Private Sub CmdSair_Click()
 wok = 1
 FrameRipley.Visible = False
End Sub

'Agregar
Private Sub Command1_Click()

If TxtAlmacen.text = "" Then
   MsgBox "Ingrese codigo del almacen ", vbInformation, "Información"
   Exit Sub
End If
VGSeleccion = 1
FrmCreacionSal.Caption = "Ingreso de Articulos"
buscar
FrmCreacionSal.Show 1
End Sub
'Modificar
Private Sub Command2_Click()
If MSFlexGrid1.Rows = 1 Then
   MsgBox "No existe registros para Modificar", vbInformation, "Información"
   Exit Sub
End If
If VGGuiaSal Then
   VGSeleccion = 2
   FrmCreacionSal.Caption = "Modificación de Articulos"
   buscar
   FrmCreacionSal.Show 1
Else
   MSFlexGrid1_Click
End If
End Sub
'Eliminar
Private Sub Command3_Click()
Dim I As Integer
If MSFlexGrid1.Rows = 1 Then
    MsgBox "No existe registros para Eliminar", vbInformation, "Información"
    Exit Sub
End If
If MsgBox("Desea Eliminar el registro", vbQuestion + vbYesNo, "Información") = vbYes Then
    I = MSFlexGrid1.RowSel
    If MSFlexGrid1.Rows > 2 Then
        MSFlexGrid1.RemoveItem I
    Else
        MSFlexGrid1.Clear
        MSFlexGrid1.Rows = 1
        MSFlexGrid1.Row = 0
        visualizarFG2
        CmdGrabarDet.SetFocus
    End If
End If
End Sub

Private Sub CmdSalirDevol_Click()
'Frame2.Visible = False
Frame1.Visible = True
CmdSalir.SetFocus
End Sub

Private Sub CmdGrabarCab_Click()
 If CmbSerie.text = "" Then
   MsgBox "Seleccione la serie de la guia", vbInformation, "Aviso"
   CmbSerie.SetFocus
   Exit Sub
 End If
 If TxTransa = "" Then
   MsgBox "Ingrese el tipo de transaccion", vbExclamation, "Error"
   TxTransa.SetFocus:    Exit Sub
 ElseIf TxTransa.Enabled And TxTransa.Visible Then
   If Existe(1, TxTransa, "TABTRANSA", "TT_CODMOV", False, "S", "TT_TIPMOV") = False Then
       MsgBox "La Transacción no existe", vbInformation, "Información"
       TxTransa.SetFocus:   Exit Sub
   End If
 End If
 If TxTransa = "TD" Then
   If Trim(TxtAlmacen) = "" Then
      MsgBox "Debe ingresar el almacen de destino", vbExclamation, "Error"
      TxtAlmacen.SetFocus
      Exit Sub
   ElseIf Existe(1, TxtAlmacen, "TABALM", "TAALMA", False) = False Then
      MsgBox "El Almacén no existe", vbInformation, "Información"
      TxtAlmacen.SetFocus: Exit Sub
   End If
 End If
 
 
'If Text1.Visible And Text1.Enabled = True Then
'    If IsNumeric(Text1.text) Then
'       If UCase(Dir$(cRuta4)) = UCase(cNomBd4) Then
'           If Existe(3, Text1, "CENTRO_COSTOS", "cencost_codigo", False) = False Then
'                  MsgBox "La Cuenta no existe", vbInformation, "Mensaje"
'                  Text1.SetFocus: Exit Sub
'           End If
'       End If
'    Else
'       MsgBox "Ingrese el numero de Centro de Costo", vbInformation, mensaje1
'       If Text1.Enabled Then Text1.SetFocus
'       Exit Sub
'    End If
'End If
 muestra
' Frame4.Enabled = True
' Txtarticulo.Enabled = True
' Txtarticulo.SetFocus
 If salir Then Unload Me
End Sub

Private Sub CmdGrabarDet_Click()
On Error GoTo GrabErr1
Dim contador As Integer               ' Grabar
Dim criterio As String
Dim cadena As String
Dim cadena1 As String
Dim cadena2 As String
Dim Aux As String
Dim rpta As Integer
Dim FACTOR As Double
Dim uSql As String
Dim tipo As String * 2
Dim nroguia As String
Dim veces As Integer, I As Integer
Dim cad, ncad As String
Dim rst As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim Productos As String
Dim ok As Integer

Set Conex = VGCNx
If Text3.text = "GR" And Text4.text = "" Then
   MsgBox "ATENCION !!! " & Chr(13) & "El Numero de Guias es 0 ", vbCritical, "Sistemas"
   Exit Sub
End If
ok = 0
CANTIDAD = 0:  veces = 0: Productos = ""
'------------------------------------------------------------------------------------------------------
'VALIDACION DE EMISION DE GUIAS
With MSFlexGrid1
    If .Rows > 1 Then
    For I = 1 To .Rows - 1
        SQL = " select stcodigo as Codigo,c.adescri as Producto,a.stskdis as Disponible,"
        SQL = SQL & .TextMatrix(I, 3) & " as Can_Pedida "
        SQL = SQL & " from dbo.stkart a inner join dbo.maeart c on stcodigo=c.acodigo "
        SQL = SQL & " where stalma='" & TxtAlmacen & "' and stcodigo='" & Trim(.TextMatrix(I, 0)) & "'"
        SQL = SQL & " And stskdis - " & .TextMatrix(I, 3) & " < 0"
        Set rs = VGCNx.Execute(SQL)
        If Not rs.EOF Then
            GridP.DataSource = rs
            ok = 1
            With GridP
                  .Columns(0).Width = 1000
                  .Columns(1).Width = 4000
                  .Columns(2).Width = 900
            End With
            Exit For
        End If
    Next I
    End If
    If ok = 1 Then
       GridP.Refresh
       MsgBox "ATENCION Saldos negativos !!! Codigo " & Trim(.TextMatrix(I, 0)) & Chr(13) & " Saldo " & rs!disponible & " NO SE PUEDE EMITIR LA GUIA ", vbCritical, "Sistemas"
       FrmValida.Visible = True
       Timer1.Enabled = True
       Exit Sub
    End If
End With
'-----------------------------------------------------------------------------------------------------
If Len(Trim(TxtAlmacen.text)) = 0 Then
    MsgBox "Falta seleccionar almacen.", vbInformation, "Sistema"
    TxtAlmacen.SetFocus
    Exit Sub
End If

'Desea grabar el registro
Set rst = Nothing
 
If MSFlexGrid1.Rows = 1 Then Exit Sub
    If Not VGGuiaSal Then
           Aux = Text4
           CmbSerie.AddItem Mid(Text4, 1, 3)
           CmbSerie.ListIndex = 0
           actualiza_guia_dev
           muestra                          'obtiene el numero de guia de salida
    Else
           Aux = Text13
           nroguia = CmbSerie.text & Text4
           'En este opción permite cambiar por número de Guia
              Do While True
                 If verifica_nro_guia(nroguia) Then          'Si es verdadero es que no esiste el nro guia
                    rpta = MsgBox("Es el número correcto de la Guia " & Chr(13) & nroguia, vbInformation + vbOKCancel, "Confirmación")
                    If rpta = vbCancel Then
                       rpta = MsgBox("desea Continuar ", vbInformation + vbOKCancel, "Mensaje")
                       If rpta = vbCancel Then
                              Exit Sub
                       End If
                    End If
                  Else
                     Exit Do
                 End If
             Loop
    End If
    
    AlmacenRF = IIf(TxtAlmacen <> "", TxtAlmacen, "")
    cTransa = TxTransa
    grabacabecera
    FACTOR = 1
    contador = 1
    'If hubo_error Then Exit Sub
    'Falta controlar si se pudo grabar puede estar bloqueada  on error goto
    While MSFlexGrid1.Rows > contador
            CANTIDAD = 0
                   
            If UCase(MSFlexGrid1.TextMatrix(contador, 0)) <> "TEXTO" Then
                  If Not VGGuiaSal Then
                         'caso de devoluciones
                         If MSFlexGrid1.TextMatrix(contador, 5) <> "" Then
                                CANTIDAD = Val(MSFlexGrid1.TextMatrix(contador, 5))
                                cantidadDEV = Val(MSFlexGrid1.TextMatrix(contador, 4))
                         Else
                                CANTIDAD = Val(MSFlexGrid1.TextMatrix(contador, 3))
                                cantidadDEV = 0  'indica no hay devolucion en ese item para que no descargue
                         End If
                  Else
                         'caso de guia remision
                         CANTIDAD = Val(MSFlexGrid1.TextMatrix(contador, 3))
                         cantidadDEV = 1
                  End If
                  If MSFlexGrid1.TextMatrix(contador, 7) = "S" Then
                  '     Data2.Recordset("DESERIE") = MSFlexGrid1.TextMatrix(contador, 2)
                           ncad = "INSERT INTO movalmdet " & _
                                  "(DEALMA,DETD,DENUMDOC,DEITEM,DECODIGO,DEDESCRI,DEUNIDAD,DECANTID,DECANTENT,DECENCOS,DEORDFAB,DEQUIPO,DESERIE,DECANREF1)" & _
                                  " VALUES (" & _
                                  "'" & TxtAlmacen.text & "'," & _
                                  "'NS','" & numserie & "'," & contador & "," & _
                                  "'" & MSFlexGrid1.TextMatrix(contador, 0) & "'," & _
                                  "'" & IIf(MSFlexGrid1.TextMatrix(contador, 1) = "", " ", MSFlexGrid1.TextMatrix(contador, 1)) & "'," & _
                                  "'" & IIf(MSFlexGrid1.TextMatrix(contador, 4) = "", " ", MSFlexGrid1.TextMatrix(contador, 4)) & "'," & CANTIDAD & "," & _
                                  CANTIDAD & ",'" & Trim(MSFlexGrid1.TextMatrix(contador, 8)) & "'," & _
                                  "'" & Trim(MSFlexGrid1.TextMatrix(contador, 9)) & "'," & _
                                  "'" & Trim(MSFlexGrid1.TextMatrix(contador, 10)) & "'," & _
                                  "'" & MSFlexGrid1.TextMatrix(contador, 2) & "'," & _
                                  MSFlexGrid1.TextMatrix(contador, 11) & ")"
                  
                  ElseIf MSFlexGrid1.TextMatrix(contador, 7) = "N" Then
                  '    Data2.Recordset("DELOTE") = MSFlexGrid1.TextMatrix(contador, 2)
                       ncad = "INSERT INTO movalmdet " & _
                                  "(DEALMA,DETD,DENUMDOC,DEITEM,DECODIGO,DEDESCRI,DEUNIDAD,DECANTID,DECANTENT,DECENCOS,DEORDFAB,DEQUIPO,DELOTE,DECANREF1)" & _
                                  " VALUES (" & _
                                  "'" & TxtAlmacen.text & "'," & _
                                  "'NS','" & numserie & "'," & contador & "," & _
                                  "'" & MSFlexGrid1.TextMatrix(contador, 0) & "'," & _
                                  "'" & IIf(MSFlexGrid1.TextMatrix(contador, 1) = "", " ", MSFlexGrid1.TextMatrix(contador, 1)) & "'," & _
                                  "'" & IIf(MSFlexGrid1.TextMatrix(contador, 4) = "", " ", MSFlexGrid1.TextMatrix(contador, 4)) & "'," & CANTIDAD & "," & _
                                  CANTIDAD & ",'" & Trim(MSFlexGrid1.TextMatrix(contador, 8)) & "'," & _
                                  "'" & Trim(MSFlexGrid1.TextMatrix(contador, 9)) & "'," & _
                                  "'" & Trim(MSFlexGrid1.TextMatrix(contador, 10)) & "'," & _
                                  "'" & MSFlexGrid1.TextMatrix(contador, 2) & "'," & _
                                  MSFlexGrid1.TextMatrix(contador, 11) & ")"
                  
                  Else
                        ncad = "INSERT INTO movalmdet " & _
                                      "(DEALMA,DETD,DENUMDOC,DEITEM,DECODIGO,DEDESCRI,DEUNIDAD,DECANTID,DECANTENT,DECENCOS,DEORDFAB,DEQUIPO,DECANREF1)" & _
                                      " VALUES (" & _
                                      "'" & TxtAlmacen.text & "'," & _
                                      "'NS','" & numserie & "'," & contador & "," & _
                                      "'" & MSFlexGrid1.TextMatrix(contador, 0) & "'," & _
                                      "'" & IIf(MSFlexGrid1.TextMatrix(contador, 1) = "", " ", MSFlexGrid1.TextMatrix(contador, 1)) & "'," & _
                                      "'" & IIf(MSFlexGrid1.TextMatrix(contador, 4) = "", " ", MSFlexGrid1.TextMatrix(contador, 4)) & "'," & CANTIDAD & "," & _
                                      CANTIDAD & ",'" & Trim(MSFlexGrid1.TextMatrix(contador, 8)) & "'," & _
                                      "'" & Trim(MSFlexGrid1.TextMatrix(contador, 9)) & "'," & _
                                      "'" & Trim(MSFlexGrid1.TextMatrix(contador, 10)) & "'," & _
                                      MSFlexGrid1.TextMatrix(contador, 11) & ")"
                  End If
                
                

                  If TxTransa <> "GF" Then   'Para descargar
                          If Not IsNull(MSFlexGrid1.TextMatrix(contador, 3)) And cantidadDEV <> 0 Then
                                grabastk (contador)
                                If (IsNumeric(MSFlexGrid1.TextMatrix(contador, 6))) > 0 And VGGuiaSal Then
                                       'SE ESTA TOMANDO DEPRECIO COMO COSTO DE PRODUCTO
                                       'Data2.Recordset("DEPRECIO") = Val(MSFlexGrid1.TextMatrix(contador, 6)) * VGTipCamb '******el precio
                                       'Data2.Recordset("DEPRECI1") = Val(MSFlexGrid1.TextMatrix(contador, 5)) * VGTipCamb '******el precio
                                     VGCNx.Execute "INSERT INTO movalmdet " & _
                                                        "(DEALMA,DETD,DENUMDOC,DEITEM,DECODIGO,DEDESCRI,DEUNIDAD,DECANTID,DECANTENT,DECENCOS,DEORDFAB,DEQUIPO," & _
                                                        "DEPRECIO,DEPRECI1,DECANREF1)" & _
                                                        " VALUES (" & _
                                                        "'" & TxtAlmacen.text & "'," & _
                                                        "'NS','" & numserie & "'," & contador & "," & _
                                                        "'" & MSFlexGrid1.TextMatrix(contador, 0) & "'," & _
                                                        "'" & IIf(MSFlexGrid1.TextMatrix(contador, 1) = "", " ", MSFlexGrid1.TextMatrix(contador, 1)) & "'," & _
                                                        "'" & IIf(MSFlexGrid1.TextMatrix(contador, 4) = "", " ", MSFlexGrid1.TextMatrix(contador, 4)) & "'," & CANTIDAD & "," & _
                                                        CANTIDAD & ",'" & Trim(MSFlexGrid1.TextMatrix(contador, 8)) & "'," & _
                                                        "'" & Trim(MSFlexGrid1.TextMatrix(contador, 9)) & "'," & _
                                                        "'" & Trim(MSFlexGrid1.TextMatrix(contador, 10)) & "'," & _
                                                        Val(MSFlexGrid1.TextMatrix(contador, 6)) * VGTipCamb & "," & _
                                                        Val(MSFlexGrid1.TextMatrix(contador, 5)) * VGTipCamb & "," & _
                                                        Val(MSFlexGrid1.TextMatrix(contador, 11)) & ")"
                                Else
'                                       Data2.Recordset("DEPRECIO") = Val(Format(precioprom, "#####.000000"))
'                                       Data2.Recordset("DEPRECI1") = IIf(IsNumeric(MSFlexGrid1.TextMatrix(contador, 5)), Val(MSFlexGrid1.TextMatrix(contador, 5)) * VGTipCamb, 0)
                                     VGCNx.Execute "INSERT INTO movalmdet " & _
                                                        "(DEALMA,DETD,DENUMDOC,DEITEM,DECODIGO,DEDESCRI,DEUNIDAD,DECANTID,DECANTENT,DECENCOS,DEORDFAB,DEQUIPO," & _
                                                        "DEPRECIO,DEPRECI1,DECANREF1)" & _
                                                        " VALUES (" & _
                                                        "'" & TxtAlmacen.text & "'," & _
                                                        "'NS','" & numserie & "'," & contador & "," & _
                                                        "'" & MSFlexGrid1.TextMatrix(contador, 0) & "'," & _
                                                        "'" & IIf(MSFlexGrid1.TextMatrix(contador, 1) = "", " ", MSFlexGrid1.TextMatrix(contador, 1)) & "'," & _
                                                        "'" & IIf(MSFlexGrid1.TextMatrix(contador, 4) = "", " ", MSFlexGrid1.TextMatrix(contador, 4)) & "'," & CANTIDAD & "," & _
                                                        CANTIDAD & ",'" & Trim(MSFlexGrid1.TextMatrix(contador, 8)) & "'," & _
                                                        "'" & Trim(MSFlexGrid1.TextMatrix(contador, 9)) & "'," & _
                                                        "'" & Trim(MSFlexGrid1.TextMatrix(contador, 10)) & "'," & _
                                                        Val(Format(precioprom, "#####.000000")) & "," & _
                                                        IIf(IsNumeric(MSFlexGrid1.TextMatrix(contador, 5)), Val(MSFlexGrid1.TextMatrix(contador, 5)) * VGTipCamb, 0) & "," & _
                                                        Val(MSFlexGrid1.TextMatrix(contador, 11)) & ")"
                                
                                End If
                                
                                
                                'En devolucion no entra las guia de transferecia
                                'Data2.Recordset.Update
                                If TxtAlmacen.text <> "" And TxTransa = "TD" And VGGuiaSal Then
                                      ''If Not IsNull(Data3.Recordset("STKPREPRO")) Then
                                      ''  precioprom = Val(Format(Data3.Recordset("STKPREPRO"), "#####.000000"))
                                      ''Else
                                        precioprom = 0
                                      ''End If
                                      Unid = ""
                                      ''If Not Data2.Recordset.EOF Then
                                      ''   If Not IsNull(Data2.Recordset("deunidad")) Then Unid = Trim(cNull(Data2.Recordset("deunidad")))
                                      ''End If
                                                                            
                                      'grabadetalmacen
                                      cad = insertar1
                                      'Completo = False
                                      Conex.Execute cad
                                      'Do
                                      '   DoEvents
                                      'Loop Until Completo
                                      grabastk1 (contador)
                               End If
                          End If
                  Else
                      VGCNx.Execute ncad
                      
'                      Data2.Recordset.Update
                      'En caso de Guia contra factura no realiza descarga a stock
 '                     TxtAlmacen = TxtAlmacen.text
                      Campo = "GS"   'El valor del campo cambia a NI u otro valor cuando es transferencia
                      nument = numserie
                      Unid = MSFlexGrid1.TextMatrix(contador, 4)
                      'cad = insertar1
                      'Completo = False
                      'Conex.Execute cad
                      'Do
                      '     DoEvents
                      'Loop Until Completo
                      'grabadetalmacen
                  End If
            Else
                Conex.Execute "INSERT INTO movalmdet " & _
                              "(DEALMA,DETD,DENUMDOC,DEITEM,DECODIGO,DEDESCRI,DEUNIDAD,DECANREF1)" & _
                              " VALUES (" & _
                              "'" & TxtAlmacen.text & "'," & _
                              "'GS','" & numserie & "'," & contador & "," & _
                              "'" & MSFlexGrid1.TextMatrix(contador, 0) & "'," & _
                              "'" & IIf(MSFlexGrid1.TextMatrix(contador, 1) = "", " ", MSFlexGrid1.TextMatrix(contador, 1)) & "'," & _
                              "'" & IIf(MSFlexGrid1.TextMatrix(contador, 4) = "", " ", MSFlexGrid1.TextMatrix(contador, 4)) & "'," & _
                              Val(MSFlexGrid1.TextMatrix(contador, 11)) & ")"
                  
'                  Data2.Recordset.Update
            End If
            contador = contador + 1
            'antes habia update
   Wend
   
If Text3.text <> "GR" Then
   Set rst = VGCNx.Execute("select * from vt_puntovtadocumento where puntovtacodigo='" & VGParametros.puntovta & "' and empresacodigo='" & VGParametros.empresacodigo & "' and documentocodigo='" & Trim(Text3) & "' and puntovtadocserie='" & CmbSerie.text & "'")
   If rst.RecordCount > 0 Then
'      VGCNx.Execute "UPDATE vt_puntovtadocumento " & _
'              " Set puntovtadoccorr='" & Right("0000000000" & Trim(CStr(Val(Text4)) + 1), 8) & "'" & _
'              " , tipo='TR' , numero='" & numserie & "'" & _
'              " where puntovtacodigo='" & VGparametros.puntovta & "' and empresacodigo='" & VGparametros.empresacodigo & "' and documentocodigo='" & Trim(Text3) & "' and puntovtadocserie='" & CmbSerie.text & "'"
   End If
   rst.Close
End If
   
'Ctr_AyudaEmpresa.xclave
If Requerimiento Then
    VGCNx.Execute "update co_cabordcompra set estadooccodigo='5' where oc_cnumord='" & Text8.text & "' " _
    & " and tipoordencodigo='" & Texttipdoc.text & "' and empresacodigo=(select empresacodigo from co_multiempresas where empresaruc='" & Text5.text & "') " '
End If
 
Requerimiento = False

If VGGuiaSal Then
     reinicia
Else
     CmdGrabarDet.Visible = False
     limpia
     visualizarFG
     If Text8.Enabled And Text8.Visible Then Text8.SetFocus
End If
VGval = False
CmdGrabarCab.Enabled = False
rpta = MsgBox("Desea Agregar Comentarios", vbYesNo + vbQuestion, "Aviso")
If rpta = vbYes Then
    'Text12.SetFocus
    crtlvisible (False)
    FrameComentario.Visible = True
    TxComentario.SetFocus
Else
     'inicializar
     TxTransa.Enabled = True
     rpta = MsgBox("¿Desea Imprimir?", vbYesNo + vbQuestion, "Aviso")
     If rpta = vbYes Then
        ' imprimir
        imprimirguias
     End If
End If
If Not VGGuiaSal Then Unload Me

Exit Sub
GrabErr1:
MsgBox Err.Number & "-" & Err.Description
Exit Sub
Resume
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
Dim numero As String
Dim distrito As String


ntabla = "movalmdet"
contador = 0

'---------------------- OPCION DE IMRPIMRI GUIAS ------------------------------------
                                   
Screen.MousePointer = 11
                                   
With MDIPrincipal.CryRptProc
        Call PropCrystal(MDIPrincipal.CryRptProc)
        If VGParametros.multiguias Then
           .ReportFileName = VGParamSistem.RutaReport & "al_guiaimpresa_" & VGParamSistem.BDEmpresa & VGParametros.empresacodigo & ".rpt"
        Else
           .ReportFileName = VGParamSistem.RutaReport & "al_guiaimpresa.rpt"
       End If
       If VGsql = 1 Then
           .Connect = "Provider=SQLOLEDB;Password=" & VGParamSistem.PwdGEN & ";Persist Security Info=True;User ID=" & VGParamSistem.UsuarioGEN & ";Initial Catalog=" & VGParamSistem.BDEmpresaGEN & ";SERVER=" & VGParamSistem.ServidorGEN
          Else
           .Connect = VGCadenaReport2
        End If
        .WindowTitle = "Impresion Guia de Remision"
        .StoredProcParam(0) = VGParamSistem.BDEmpresa
        .StoredProcParam(1) = TxtAlmacen.text
        .StoredProcParam(2) = "GR"
        .StoredProcParam(3) = nroguia
        .Action = 1
  End With
Screen.MousePointer = 1
Exit Sub
errores:
  Screen.MousePointer = 1
  MsgBox "Error inesperado: " & Err.Number & "  " & Err.Description, vbExclamation
  Err = 0
  Exit Sub
  
End Sub


Private Sub CmdSalir_Click()
Dim I As Integer
Dim Productos As String
Dim ok As Integer
ok = 0
If Frame1.Visible Then
     If MSFlexGrid1.Rows > 1 Then
        If vbYes = MsgBox("Desea Grabar?", vbYesNo + vbQuestion, "Aviso") Then
           With MSFlexGrid1
           For I = 1 To .Rows - 1
              SQL = " select stcodigo as Codigo,c.adescri as Producto,a.stskdis as Disponible,"
              SQL = SQL & .TextMatrix(I, 3) & " as Can_Pedida "
              SQL = SQL & " from dbo.stkart a inner join dbo.maeart c on stcodigo=c.acodigo "
              SQL = SQL & " where stalma='" & TxtAlmacen & "' and stcodigo='" & Trim(.TextMatrix(I, 0)) & "'"
              SQL = SQL & " And stskdis - " & .TextMatrix(I, 3) & " < 0"
              Set rs = VGCNx.Execute(SQL)
              If Not rs.EOF Then
                 GridP.DataSource = rs
                 ok = 1
                 With GridP
                    .Columns(0).Width = 1000
                    .Columns(1).Width = 4000
                    .Columns(2).Width = 900
                 End With
                 Exit For
              End If
           Next I
        If ok = 1 Then
           GridP.Refresh
           MsgBox "ATENCION Saldos negativos !!! Codigo " & Trim(.TextMatrix(I, 0)) & Chr(13) & " Saldo " & rs!disponible & " NO SE PUEDE EMITIR LA GUIA ", vbCritical, "Sistemas"
           FrmValida.Visible = True
           Timer1.Enabled = True
           Exit Sub
        Else
           CmdGrabarDet_Click
        End If
        End With
      End If
           VGval = False
           TxTransa.Enabled = True
           Text6.Enabled = True
           Text5.Enabled = True
           Text8.Enabled = True
           reinicia
    End If
    Unload Me
 Else
    Frame1.Visible = True
    CmdSalir.SetFocus
 End If

End Sub
Private Sub Command4_Click()
FrmValida.Visible = False
Timer1.Enabled = False
t = 0
Text3.SetFocus
End Sub

Private Sub Conex_ExecuteComplete(ByVal RecordsAffected As Long, ByVal pError As ADODB.error, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.Recordset, ByVal pConnection As ADODB.Connection)
  Completo = True
End Sub

Private Sub Ctr_AyudaEmpresa_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
Set Rs2 = Nothing
Set Rs2 = VGCNx.Execute("select empresacodigo as Empresa,tipo as Tipo,pedidotipofac as Doc,PEDIDOENTREGA AS Destino ,[Nro Orden]=PedidoNumero," _
& " (detpedcantpedida) as Cantidad from v_almacenyventas WHERE TIPO=3  and empresacodigo='" & ColecCampos(0) & "'" _
& " ")
If Not Rs2.EOF Then
    Set TDBGrid.DataSource = Rs2
    With TDBGrid
        .Columns(0).Width = 700
        .Columns(1).Width = 500
        .Columns(2).Width = 500
        .Columns(3).Width = 3500
        .Columns(4).Width = 1200
        .Columns(5).Width = 1000
    End With
    
    TDBGrid.Refresh
    FrmPen.Visible = True
    Text5.text = ColecCampos(1)
    Text6.text = ColecCampos(2)
    Text7.text = ColecCampos(3)
Else
    MsgBox "No hay Requerimientos pendientes" & Chr(13) & "para esta empresa.", vbInformation, "Sistemas"
    Ctr_AyudaEmpresa.SetFocus
    Text5.text = ColecCampos(1)
    Text6.text = ColecCampos(2)
    Text7.text = ESNULO(ColecCampos(3), "")
    'Exit Sub
End If

'Text5.text = ColecCampos(1)
'Text6.text = ColecCampos(2)
'Text7.text = ColecCampos(3)

CmdGrabarDet.Visible = True
Command3.Visible = True
Command2.Visible = True
Command1.Visible = True



End Sub

Private Sub Ctr_AyuTransporte_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
TxtTransp.text = Ctr_AyuTransporte.xclave
End Sub

Private Sub Ctr_AyuVendedor_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
'TxVendedor.text = Ctr_AyuVendedor.xclave
End Sub

Private Sub DTPicker1_Change()
        DTPicker1.Value = UltimoCierreFech(DTPicker1.Value)
        VGTipCamb = DevolverTCambio(DTPicker1.Value)
End Sub

Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
  If TxTransa = "TD" Then
         TxtAlmacen.SetFocus
  Else
         If TxTransa.Enabled Then
            TxTransa.SetFocus
         End If
  End If
End If
End Sub

Private Sub Form_Activate()
   Dim J, kTotal As Double
   If MSFlexGrid1.Rows > 1 Then
      Text2 = Format(MSFlexGrid1.Rows - 1, "##,###,##0.00")
      kTotal = 0
      For J = 1 To MSFlexGrid1.Rows - 1
        kTotal = kTotal + CDbl(MSFlexGrid1.TextMatrix(J, 3))
      Next
      Text9 = Format(kTotal, "##,###,##0.00")
   Else
      Text2 = Format(0, "##,###,##0.00")
      Text9 = Format(0, "##,###,##0.00")
   End If
End Sub

Private Sub Form_Load()
Dim rsqli As String
Call Ctr_AyuTransporte.conexion(VGCNx)
Call Ctr_AyuVendedor.conexion(VGCNx)
Call Ctr_AyudaEmpresa.conexion(VGCNx)
Cliente = False
Requerimiento = False

VGSeleccion = 1   'Indica el modo de apertura = 1 y modificacion=2
VGForm = 6
limpia

Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = 800
 

salir = False
hubo_error = False
 DTPicker1.MaxDate = VGParamSistem.fechatrabajo
 DTPicker1.Value = UltimoCierreFech(CDate(Format(Now, "dd/MM/yyyy")))
VGTipCamb = DevolverTCambio(DTPicker1.Value)
'SAS
Deshabilitar (False)
Codigo2 = "GUIA DE REMISION"
If VGGuiaSal Then
    FrmGuiaSal.Caption = "Registro de Guia de Salidas"
    Command1.Visible = False
    Command2.Visible = False
    Command3.Visible = False
    'CmdGrabarDet.Visible = False
    lblSerie.Visible = True
    AgregarSerie
    CmbSerie.Visible = CBool(-1)
    visualizarFG1
Else
    FrmGuiaSal.Caption = "Devolucion de Guia de Salidas"
    visualizarFG
    TxTransa.Enabled = False
    Text3.Enabled = False
    Text4.Enabled = True
    TxTransa = "DG"
    Text3 = "GS"
    Command1.Visible = False
    Command2.Visible = False
    Command3.Visible = False
    'CmdGrabarDet.Visible = False
    CmdGrabarCab.Visible = False
    CmbSerie.Visible = False
End If


If CmbSerie.Visible Then CmbSerie.ListIndex = 0

CmdGrabarDet.Picture = MDIPrincipal.ImageList2.ListImages("Facturado").Picture
CmdSalir.Picture = MDIPrincipal.ImageList2.ListImages("Retornar").Picture
Command3.Picture = MDIPrincipal.ImageList2.ListImages("Eliminar").Picture
Command2.Picture = MDIPrincipal.ImageList2.ListImages("Modificar").Picture
Command1.Picture = MDIPrincipal.ImageList2.ListImages("Insertar").Picture

End Sub

Private Sub Grid1_DblClick()
MSFlexGrid1.Rows = 1


Set rs = VGCNx.Execute("select A.PEDIDONUMERO,A.PEDIDONROFACT, b.PRODUCTOCODIGO,c.ADESCRI," _
& " B.DETPEDCANTPEDIDA,A.PEDIDOTOTNETO," _
& " from VT_PEDIDO a inner join VT_DETALLEPEDIDO b  on a.PEDIDONUMERO=b.PEDIDONUMERO " _
& " Inner join maeart c on b.PRODUCTOCODIGO=c.acodigo " _
& " where A.PEDIDONROFACT='" & Rs2!pedidonumero & "'")

Do While Not rs.EOF
    MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
    MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 0) = rs!productocodigo
    MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 1) = rs!adescri
    MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 3) = rs!detpedcantpedida
    MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 5) = rs!pedidototneto
    rs.MoveNext
Loop

FrmPen.Visible = False

End Sub


Private Sub MSFlexGrid1_Click()
If Not VGGuiaSal Then
    Frame1.Visible = False
'    Frame2.Visible = True
    llenadatos
 End If
End Sub

Private Sub TDBGrid_DblClick()
If Rs2(0) <> "3" Then
    Set rs = VGCNx.Execute("select pedidonumero,pedidotipofac,pedidonrofact,productocodigo,ADESCRI," _
    & " Saldo=sum(detpedcantpedida)-sum(isnull(decantid,0)),PEDIDOTOTNETO=0,VENDEDORCODIGO," _
    & " almacencodigo , transportecodigo,pedidoentrega,pedidoobserva " _
    & " from " & VGParamSistem.BDEmpresa & ".dbo.v_almacenyventas  " _
    & " WHERE clienteruc='" & ruc & "' and empresacodigo='" & VGParametros.empresacodigo & "' and pedidonumero='" & Rs2(3) & "' " _
    & " group by pedidonumero,pedidotipofac,pedidonrofact,productocodigo,ADESCRI,VENDEDORCODIGO," _
    & " almacencodigo ,transportecodigo,pedidoentrega,pedidoobserva having (sum(detpedcantpedida))-sum(isnull(decantid,0))>0 ")
tipofactura = rs!pedidotipofac
nroguia = rs!pedidonrofact
Else
   Set rs = VGCNx.Execute("SELECT a.tipoordencodigo,b.oc_ccodigo,dbo.MAEART.ADESCRI,b.oc_ncantid,a.oc_cnumord as NroPedido," _
   & " a.oc_cnumord,a.almacenorigen FROM dbo.co_DETordcompra b " _
   & " inner JOIN dbo.co_cabordcompra a on b.OC_CNUMORD=a.OC_CNUMORD " _
   & " inner JOIN dbo.MAEART ON  b.OC_Ccodigo=maeart.acodigo inner JOIN dbo.tabalm c " _
   & " ON  a.almacendestino=c.taalma where a.estadooccodigo<=4 and a.oc_cnumord='" & Rs2(3) & "'")
tipofactura = ""
nroguia = ""
End If
MSFlexGrid1.Rows = 1
    rs.MoveFirst
If Rs2(0) = 3 Then
    Do While Not rs.EOF
        MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
        MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 0) = rs!oc_ccodigo
        MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 1) = rs!adescri
        MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 3) = rs!oc_ncantid
        MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 5) = 0
        Txtnrodoc.text = rs!oc_cnumord
        Texttipdoc.text = rs!tipoordencodigo
        Text8.text = rs!oc_cnumord
        TxtAlmacen.text = rs!almacenorigen
        rs.MoveNext
    Loop
    Requerimiento = True
Else

    Do While Not rs.EOF
        MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
        MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 0) = RTrim(rs!productocodigo)
        MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 1) = rs!adescri
        MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 3) = rs!saldo
        MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 5) = rs!pedidototneto
        Txtnrodoc.text = rs!pedidonumero
        TxtAlmacen.text = rs!almacencodigo
        If Not IsNull(rs!vendedorcodigo) Then
            Ctr_AyuVendedor.xclave = rs!vendedorcodigo: Ctr_AyuVendedor.Ejecutar
        End If
        
        If Not IsNull(rs!transportecodigo) Then
            Ctr_AyuTransporte.xclave = rs!transportecodigo: Ctr_AyuTransporte.Ejecutar
        End If
        Text7.text = Trim(Escadena(rs!pedidoentrega))
        TxtCon.text = rs!pedidoobserva
        rs.MoveNext
    Loop
    Requerimiento = False
End If

FrmPen.Visible = False

End Sub
'Almacen
Private Sub TxtAlmacen_DblClick()
Dim Adodc3 As New ADODB.Recordset
    Set Adodc3 = VGCNx.Execute("SELECT TAALMA,TADESCRI FROM TABALM")
    frmReferencia.Conectar Adodc3, "SELECT TAALMA,TADESCRI FROM TABALM"
    frmReferencia.Label1.Caption = "Almacenes"
    frmReferencia.Show vbModal
    If vGUtil(1) <> "" Then TxtAlmacen = (vGUtil(1))
    If TxtAlmacen <> "" Then TxtAlmacen_KeyPress (13)
    VGAlma = TxtAlmacen
    If VGRegEnt = 2 Then
     FrmCreacionSal.Ctr_Ayuart.filtro = " stalma='" & VGAlma & "' and stskdis> 0  "
  Else
  FrmCreacionSal.Ctr_Ayuart.filtro = ""
  End If
    
End Sub

Private Sub TxtAlmacen_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then TxtAlmacen_DblClick
End Sub

Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
 Dim rst As New ADODB.Recordset
  
 If KeyCode = 112 Then
    Text3_DblClick
 ElseIf KeyCode = 13 Then
   If UCase(Text3.text) = "GR" Then
   SQL = "select * from vt_puntovtadocumento where empresacodigo='" & VGParametros.empresacodigo & "' and puntovtacodigo='" & VGParametros.puntovta & "' and documentocodigo='" & Trim(Text3) & "'"
      Set rst = VGCNx.Execute(SQL)
      If rst.RecordCount > 0 Then
         CmbSerie.Clear
         Do Until rst.EOF
            CmbSerie.AddItem rst!puntovtadocserie
            Text4.text = Format(Trim(rst!puntovtadoccorr), "0000000000")
            rst.MoveNext
         Loop
         CmbSerie.ListIndex = 0
      
      End If
      rst.Close
    End If
 End If
 
 Set rst = Nothing
End Sub

Private Sub Text3_LostFocus()
   On Error Resume Next
   Call Text3_KeyDown(13, 0)
End Sub

Private Sub Text4_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 112 And Text3 <> "" Then
    Text4_DblClick
 End If
End Sub

Private Sub Text4_LostFocus()
Text4.text = Format(Text4.text, "0000000000")
End Sub

Private Sub Text5_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
     Text5_DblClick
ElseIf KeyCode = 8 Then
     Text6 = ""
     Text7 = ""
End If
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
Dim rbusca As New ADODB.Recordset

If KeyAscii = 13 And Trim(Text5) <> "" Then
      Text5 = Trim(Text5)
           
      Text6 = existe_clie(Text5)
       If Text6 = "" Then
              MsgBox "No existe el codigo del cliente", vbExclamation, "Clientes"
       Else
             Text7 = Mid(direccion, 1, 25)
             'TxtCambio.SetFocus
       End If
End If
End Sub

Private Sub TxtAlmacen_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And TxTransa.text = "TD" And Len(TxTransa.text) = 2 Then
    If Existe(1, TxtAlmacen, "TabAlm", "TAALMA", False) Then
        SendKeys "{tab}"
    End If
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
Dim criterio As String

If KeyAscii = 13 Then
    If Len(Text3) = 2 Then
        Text3 = UCase(Text3)
        If ValidarDoc(Text3) = "" Then Exit Sub
        
        Text4.SetFocus
        Exit Sub
    Else
        Text3 = ""
        Text4 = ""
        SendKeys "{tab}"
        KeyAscii = 0
    End If
End If

End Sub

Private Sub Timer1_Timer()
t = t + 1
Lblmensaje.Visible = IIf(t Mod 2 = 0, True, False)
If t > 21 Then
    Timer1.Enabled = False
    t = 0
End If
End Sub

Private Sub TxTransa_DblClick()
Dim Adodc3 As ADODB.Recordset

Set Adodc3 = New ADODB.Recordset
VGRegEnt = 2

Set Adodc3 = VGCNx.Execute("SELECT TT_CODMOV,TT_DESCRI,tt_clie FROM Tabtransa where  TT_tipmov = '" & IIf(VGRegEnt = 1, "I", "S") & "'")
frmReferencia.Conectar Adodc3, "SELECT TT_CODMOV,TT_DESCRI,tt_clie FROM Tabtransa where  TT_tipmov = '" & IIf(VGRegEnt = 1, "I", "S") & "'"
frmReferencia.Label1.Caption = "Transacciones"
frmReferencia.Show vbModal

If vGUtil(1) <> "" Then
    TxTransa = vGUtil(1)
    Label9.Caption = vGUtil(2)
End If

If vGUtil(3) <> "" Then Cliente = IIf(vGUtil(3) = "S", True, False)

If TxTransa <> "" Then
    Deshabilitar (True)
                  
    If Cliente Then
        Text5.Enabled = True
        Text5.Visible = True
        Ctr_AyudaEmpresa.Enabled = False
        Ctr_AyudaEmpresa.Visible = False
        Ctr_AyudaEmpresa.xclave = ""
        Label12.Caption = "Cod.Cliente :"
    Else
        Ctr_AyudaEmpresa.Enabled = True
        Ctr_AyudaEmpresa.Visible = True
        Text5.Enabled = False
        Text5.Visible = False
        Label12.Caption = "Cod.Proveedor :"
    End If

    LIMPIACABECERA
    buscar
End If

End Sub
Private Sub TxTransa_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Len(TxTransa.text) = 2 Then
    TxTransa = UCase(TxTransa)
    If (transa(TxTransa) <> "") Then
            Deshabilitar (True)
            If Cliente Then
                Text5.Enabled = True
                Text5.Visible = True
                Ctr_AyudaEmpresa.Enabled = False
                Ctr_AyudaEmpresa.Visible = False
                Ctr_AyudaEmpresa.xclave = ""
                Label12.Caption = "Cod.Cliente :"
                
            Else
                Ctr_AyudaEmpresa.Enabled = True
                Ctr_AyudaEmpresa.Visible = True
                Text5.Enabled = False
                Text5.Visible = False
                Label12.Caption = "Cod.Proveedor :"
            End If
            LIMPIACABECERA
            buscar
            CmbSerie.Clear
            Call AgregarSerie
            If TxTransa = "TD" Then
                   TxtAlmacen.SetFocus
            Else
                   Text3.SetFocus
            End If
     Else
            Enfoque TxTransa
    End If
Else
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End If
End Sub
Private Sub TxTransa_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    TxTransa_DblClick
    CmbSerie.Clear
    Call AgregarSerie
End If

End Sub
Private Sub Text4_DblClick()
If Not VGGuiaSal Then
     FormAyuguia.Show 1
     If Text4 <> "" Then
            devolver (Text4)
     End If
End If
If Text4 <> "" Then
     Text4_KeyPress (13)
End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)

 If KeyAscii = 13 And TxTransa = "GF" Then
  If Text4 = "" Then
     'Text3 = "FT"
   MsgBox "Ingrese el número de factura", vbInformation, "Aviso"
   Text4.SetFocus
  End If
  If (Text3 = "FT" Or Text3 = "FE" Or Text3 = "BV") And VGGuiaSal Then '
    llenarfactura
    Exit Sub
  End If
 End If
 If KeyAscii = 13 Then
    If Not VGGuiaSal And Trim(Text4) <> "" Then
         devolver (Text4)
    Else
         SendKeys "{tab}"
         KeyAscii = 0
    End If
 End If
End Sub

Private Sub Text5_DblClick()
Dim acliente As New ADODB.Recordset
wok = 0
If Len(Trim(TxTransa.text)) = 0 Then
    MsgBox "Primero seleccione tipo de transaccion", vbInformation, "Sistema"
    TxTransa.SetFocus
    Exit Sub
End If

If Len(Text3.text) = 0 Then
    MsgBox "Primero ingrese documento de referencia", vbInformation, "Sistema"
    Text3.SetFocus
    Exit Sub
ElseIf Len(CmbSerie.text) = 0 Then
    MsgBox "Falta seleccionar serie", vbInformation, "Sistema"
    CmbSerie.SetFocus
    Exit Sub
ElseIf Len(Text4.text) = 0 Then
    MsgBox "Falta indicar numero de documento", vbInformation, "Sistema"
    Text4.SetFocus
    Exit Sub
End If

Text5 = ""
Text6 = ""
Text7 = ""
FrmAyuCliente.Show 1
Text5 = FrmAyuCliente.cCod
Text6 = FrmAyuCliente.cNom
Text7 = Trim(FrmAyuCliente.cDir)
wok = FrmAyuCliente.guiasTerceros

If FrmAyuCliente.cRuc <> "" Then
   ruc = FrmAyuCliente.cRuc
Else
   ruc = FrmAyuCliente.cCod
End If

Dim RsCliente As ADODB.Recordset

Set RsCliente = VGCNx.Execute("select empresacodigo from co_multiempresas where empresaruc='" & ruc & "'")
    If wok = 1 Then
       Call clienteripley
    End If
If analitico = 1 Then
   SQL = " clientecodigo='" & Text5 & "' and proyectocierre=0 and tipoanaliticocodigo='" & VGParamSistem.tipoanaliticocodigo & "'"
   Set acliente = VGCNx.Execute(" select * from gr_proyectos where " & SQL)
   If acliente.RecordCount = 0 Then
      MsgBox ("No existe proyectos activos para este cliente ")
      Text5.SetFocus
      FrmCreacionSal.Ctr_AyuAnalitico.Visible = False
      Exit Sub
    Else
      FrmCreacionSal.Ctr_AyuAnalitico.filtro = SQL
  End If
End If
If RsCliente.RecordCount = 0 Then
    CmdGrabarDet.Visible = True
    Command3.Visible = True
    Command2.Visible = True
    Command1.Visible = True
    SQL = " select Tipo, empresacodigo as Empresa,Almacen=almacencodigo ,Numero_Pedido=PedidoNumero,"
    SQL = SQL & " Saldo=(detpedcantpedida)-sum(isnull(decantid,0)) from " & VGParamSistem.BDEmpresa & ".dbo.v_almacenyventas"
    SQL = SQL & " WHERE CLIENTEruc='" & ruc & "' and empresacodigo='" & VGParametros.empresacodigo & "'"
    SQL = SQL & " and Estado=0 group by tipo,almacencodigo,pedidonumero,detpedcantpedida,empresacodigo "
    SQL = SQL & " having (detpedcantpedida)-sum(isnull(decantid,0))>0"
    Set Rs2 = Nothing
    Set Rs2 = VGCNx.Execute(SQL)

    If Not Rs2.EOF Then
        Set TDBGrid.DataSource = Rs2
        TDBGrid.Refresh
        FrmPen.Visible = True
    Else
   '     MsgBox "Este cliente no tiene Guias pendientes", vbInformation, "Sistemas"
    End If
Else
    CmdGrabarDet.Visible = True
    Command3.Visible = True
    Command2.Visible = True
    Command1.Visible = True
    Set Rs2 = Nothing
    Set Rs2 = VGCNx.Execute("select tipo as Tipo,pedidotipofac as Doc,PEDIDOENTREGA AS Destino ,[Nro Orden]=PedidoNumero," _
    & " (detpedcantpedida) as Cantidad from v_almacenyventas WHERE TIPO=3 ")
    If Not Rs2.EOF Then
        Set TDBGrid.DataSource = Rs2
        TDBGrid.Refresh
        FrmPen.Visible = True
    Else
        MsgBox "Este cliente no tiene Guias pendientes"
    End If
End If


End Sub
Private Sub clienteripley()
   FrameRipley.Visible = True
   
End Sub

Private Sub Text3_DblClick()
Dim Adodc3 As ADODB.Recordset

Set Adodc3 = New ADODB.Recordset

Set Adodc3 = VGCNx.Execute("SELECT TDO_TIPDOC,TDO_DESCRI  FROM TIPO_DOCU")
frmReferencia.Conectar Adodc3, "SELECT TDO_TIPDOC,TDO_DESCRI  FROM TIPO_DOCU"
frmReferencia.Label1.Caption = "Tipo de Documentos"
frmReferencia.Show vbModal


If vGUtil(1) <> "" Then Text3 = (vGUtil(1))
If vGUtil(1) <> "" Then Label10.Caption = (vGUtil(2))

If Text3 <> "" Then Text4.SetFocus
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 And Trim(Text7.text) <> "" Then
       Text7.SetFocus
  End If
End Sub


Private Sub Text8_KeyPress(KeyAscii As Integer)
  Dim criterio As String
  If KeyAscii = 13 Then             'de orden de compra
    If IsNumeric(Text8.text) Then
       'If Len(Text4.text) = 7 Then
      muestra
    Else
      If Trim(TxTransa) = "" Then
        CmdSalir.SetFocus
      End If
      MsgBox "Ingrese el numero de la Orden Compra", vbOKOnly + vbExclamation, "Error"
   End If
  End If
End Sub
Private Sub ocultarlabel()
label7.Visible = True   'False
Text7.Visible = True    'False
Label9.Visible = True   'False

Label10.Visible = True   'False
Label11.Visible = True    'False
TxtAlmacen.Visible = True    'False
End Sub

Private Sub muestra()
Dim numfil As Long
Dim RSQL As String
Dim ultimoserie As Long
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
If Trim(TxtAlmacen.text) <> "" Then
   RSQL = "Select CTNNUMERO,CTNNUMFIN FROM NUM_DOCUMENTOS WHERE   CTNCODIGO = 'GS'  AND CTNNUMSER = '" & CmbSerie.text & "' "
   rs.Open RSQL, VGCNx, adOpenDynamic, adLockOptimistic
   If IsNull(rs(0)) Then
      MsgBox "No se ha ingresado el numero de inicio de la serie en la Tabla ", vbInformation, "Error"
      salir = True
      Exit Sub
   End If
   If Not rs.EOF Then
      numsal = rs(0) + 1
      ultimoserie = rs(1)
      Serie = Format(CmbSerie.text, "000")            ' ********************* Serie contiene   la seie de la guia de remision
      If rs(0) > rs(1) Then
         MsgBox "No se puede emitir guia," & Chr(13) & "La Nro. guia es mayor que número máximo", vbCritical, "Aviso"
         salir = True
         Exit Sub
      End If
      numserie = Serie & Format(numsal, "00000000")
      If VGGuiaSal And TxTransa <> "GF" Then
         sigue
      End If
   End If
   rs.Close
Else
   MsgBox "No hay ningún Almacén Activo", vbInformation, "Información"
End If

If Not CmdGrabarDet.Visible Then sigue: CmdGrabarDet.SetFocus

End Sub


Private Sub sigue()
  Command1.Visible = True
  Command2.Visible = True
  Command3.Visible = True
  CmdGrabarDet.Visible = True
  FrmCreacionSal.Caption = "Ingreso de Articulos"
  buscar
  If ChkTalla.Value = 0 Then
    FrmCreacionSal.Show 1
   Else
    FrmIngTallas.Show 1
  End If
End Sub
Public Function insertar1()            ' grabadetalmacen()
'Esta funcion graba el detalle en el almacen de transferecia
 Dim cad As String
 
 If MSFlexGrid1.TextMatrix(contador, 7) = "S" Then
      cad = "insert into MovAlmDet (DEALMA,DETD,DENUMDOC,DECODIGO,DECANTID,DEPRECIO,DEITEM,DEUNIDAD,DESERIE,DETIPCAM,DECENCOS,DEORDFAB,DEQUIPO) values ('" & TxtAlmacen & "','" & Campo & "','" & Format(nument, "0000000000") & "','" & MSFlexGrid1.TextMatrix(contador, 0) & "'," & CANTIDAD & "," & precioprom & "," & contador & ",'" & Unid & "','" & MSFlexGrid1.TextMatrix(contador, 2) & "'," & TCamb & ",'" & MSFlexGrid1.TextMatrix(contador, 8) & "','" & MSFlexGrid1.TextMatrix(contador, 9) & "','" & MSFlexGrid1.TextMatrix(contador, 10) & "' ) "
 ElseIf MSFlexGrid1.TextMatrix(contador, 7) = "N" Then
      cad = "insert into MovAlmDet (DEALMA,DETD,DENUMDOC,DECODIGO,DECANTID,DEPRECIO,DEITEM,DEUNIDAD,DELOTE,DETIPCAM,DECENCOS,DEORDFAB,DEQUIPO) values ('" & TxtAlmacen & "','" & Campo & "','" & Format(nument, "0000000000") & "','" & MSFlexGrid1.TextMatrix(contador, 0) & "'," & CANTIDAD & "," & precioprom & "," & contador & ",'" & Unid & "','" & MSFlexGrid1.TextMatrix(contador, 2) & "' ," & TCamb & ",'" & MSFlexGrid1.TextMatrix(contador, 8) & "','" & MSFlexGrid1.TextMatrix(contador, 9) & "','" & MSFlexGrid1.TextMatrix(contador, 10) & "') "
 Else
      cad = "insert into MovAlmDet (DEALMA,DETD,DENUMDOC,DECODIGO,DECANTID,DEPRECIO,DEITEM,DEUNIDAD,DETIPCAM,DECENCOS,DEORDFAB,DEQUIPO) values ('" & TxtAlmacen & "','" & Campo & "','" & Format(nument, "0000000000") & "','" & MSFlexGrid1.TextMatrix(contador, 0) & "'," & CANTIDAD & "," & precioprom & "," & contador & ",'" & Unid & "'," & TCamb & ",'" & MSFlexGrid1.TextMatrix(contador, 8) & "','" & MSFlexGrid1.TextMatrix(contador, 9) & "','" & MSFlexGrid1.TextMatrix(contador, 10) & "') "
 End If
 insertar1 = cad
End Function

Public Sub grabaalmacen()
'GRABA EN EL ALMCEN DESTINO
Dim uSql As String
Dim insertar1 As String
Dim rs As New ADODB.Recordset
Dim RSQL As String

RSQL = "select  TANUMENT from tabAlm where TAALMA =  '" & TxtAlmacen & " ' "
Set rs = VGCNx.Execute(RSQL)
If rs.EOF Then Exit Sub
nument = rs(0): Campo = "NI"

If UCase(Dir$(cRuta4)) = UCase(cNomBd4) Then
   TCamb = Val(Devolver_Dato(3, DTPicker1, "Tipo_Cambio", "TIPOCAMB_FECHA", True, "TIPOCAMB_VENTA", "02", "TIPOMON_CODIGO"))
ElseIf UCase(Dir$(cRuta2)) = "BDCOMUN.MDB" Then
   TCamb = Val(Devolver_Dato(1, DTPicker1, "Tipo_Cambio", "TIPOCAMB_FECHA", True, "TIPOCAMB_VENTA", "02", "TIPOMON_CODIGO"))
End If
  
insertar1 = "insert into MovAlmCab (CAALMA,CATD,CANUMDOC,CACODMOV,CAFECDOC,CATIPMOV,CASITGUI,CARFALMA,CARFTDOC,CARFNDOC,CAHORA,CAUSUARI,catipcam,contacto) values ('" & TxtAlmacen & "','" & Campo & "','" & Format(nument, "00000000000") & "','51','" & DTPicker1.Value & "','I','V','" & TxtAlmacen.text & "','" & Text3 & "','" & nroguia & "','" & Format(Time, "hh:mm:ss") & "','" & VGUsuario & "'," & TCamb & ",'" & TxtCon.text & "') "
 
VGCNx.Execute insertar1
uSql = "Update TabAlm set TANUMENT = " & nument + 1 & " where TAALMA='" & TxtAlmacen & "' "
VGCNx.Execute uSql
'insertar1 = "insert into MovAlmCab (CAALMA,CATD,CANUMDOC,CACODMOV,CAFECDOC,CATIPMOV,CASITUA,CARFTDOC,CARFNDOC,CARFALMA) values ('" & TxtAlmacen & "','" & Campo & "','" & nument & "','TD','" & DTPicker1 & "','I','V','NS','" & Text4 & "','01' ) "

End Sub

Public Sub grabastk(contador)
Dim cadena As String
Dim criterio As String
Dim ncantidad As Double
Dim acmd As New ADODB.Command
Dim RSBUSCA2 As New ADODB.Recordset

   On Error GoTo GrabErr
   cadena = MSFlexGrid1.TextMatrix(contador, 0)
   criterio = " STCODIGO = '" & cadena & "'"
   criterio = criterio + " and  STALMA = '" & TxtAlmacen.text & "'"
   'Data3.Recordset.FindFirst criterio
   Set RSBUSCA2 = VGCNx.Execute("SELECT * FROM STKART WHERE " & criterio)
   If RSBUSCA2.RecordCount > 0 Then
            'Data3.Recordset.Edit
            canttemp = RSBUSCA2.Fields("STSKDIS")      ' revisar si validar en creacion
                           'para la salida
            If VGGuiaSal Then
              'Data3.Recordset("STSKDIS") = Data3.Recordset("STSKDIS") - CANTIDAD
              ncantidad = RSBUSCA2.Fields("STSKDIS") - CANTIDAD
            Else
              'Data3.Recordset("STSKDIS") = Data3.Recordset("STSKDIS") + cantidadDEV
              ncantidad = RSBUSCA2.Fields("STSKDIS") + cantidadDEV
            End If
            If Not IsNull(RSBUSCA2.Fields("STKPREPRO")) Then
              If VGGuiaSal Then
                 precioprom = RSBUSCA2.Fields("STKPREPRO") * VGTipCamb
              Else
                 precioprom = RSBUSCA2.Fields("STKPREPRO")
              End If
            Else
              precioprom = 0
            End If
            
            VGCNx.Execute "UPDATE stkart " & _
                              " set STSKDIS=" & ncantidad & _
                               " WHERE " & criterio
    Else
'            Data3.Recordset.AddNew                  'existe
 '           Data3.Recordset("STALMA") = TxtAlmacen.text   '"01"
 '           Data3.Recordset("STCODIGO") = MSFlexGrid1.TextMatrix(contador, 0)
 '           Data3.Recordset("STSKDIS") = CANTIDAD
            VGCNx.Execute "INSERT INTO stkart " & _
                            "(STALMA,STCODIGO,STSKDIS)" & _
                            " VALUES(" & _
                            "'" & TxtAlmacen.text & "'," & _
                            MSFlexGrid1.TextMatrix(contador, 0) & "," & _
                            CANTIDAD & ")"
                            
        'Grabamos en Facturacion
         Set acmd.ActiveConnection = VGCNx
         acmd.CommandText = "al_actualizaproducto_pro"
         acmd.CommandType = adCmdStoredProc
         acmd.Prepared = True
         With acmd
             .Parameters("@baseini") = VGCNx.DefaultDatabase
             .Parameters("@basefin") = VGCNx.DefaultDatabase  'VGBase2
             .Parameters("@almacen") = TxtAlmacen.text
             .Parameters("@articulo") = MSFlexGrid1.TextMatrix(contador, 0)
             .Parameters("@tipo") = "1"
         End With
         acmd.Execute
         Set acmd = Nothing
    End If
    'Data3.Recordset.Update
    RSBUSCA2.Close
    Set RSBUSCA2 = Nothing
     
     
    If MSFlexGrid1.TextMatrix(contador, 7) = "S" Then grabaserie TxtAlmacen.text, MSFlexGrid1.TextMatrix(contador, 0)
    If MSFlexGrid1.TextMatrix(contador, 7) = "N" Then grabalote TxtAlmacen.text, MSFlexGrid1.TextMatrix(contador, 0)
    Call ValMes(TxtAlmacen.text, False)
    Exit Sub
GrabErr:
'Resume Next
    MsgBox Err.Description
    hubo_error = True
End Sub

Private Sub grabalote(alma As String, codigo As String)
Dim uSql As String
Dim Lote As String
Dim nuevo_stk As Double
Dim RSQL As String
Dim rs As New ADODB.Recordset
Dim fecfab As Date
Dim fecven As Date
   
    Lote = MSFlexGrid1.TextMatrix(contador, 2)
    RSQL = "select STSLKDIS FROM STKLOTE where   STSALMA= '" & alma & "' and  STSCODIGO= '" & codigo & "' and STSLOTE= '" & Lote & "'" '
    'Set RS = VGBaseDatos.OpenRecordset(RSQL, dbOpenSnapshot)
    Set rs = VGCNx.Execute(RSQL)
    If rs.RecordCount > 0 Then
       If (Campo = "NI" And alma <> TxtAlmacen.text) Then
         nuevo_stk = rs(0) + CANTIDAD
       Else
         nuevo_stk = rs(0) - CANTIDAD
       End If
       uSql = "Update STKLOTE set STSLKDIS = " & nuevo_stk & " WHERE  STSALMA='" & alma & "' and STSCODIGO='" & codigo & "'AND STSLOTE='" & Lote & "'"
    Else
     If (MSFlexGrid1.TextMatrix(contador, 9)) <> "__/__/____" And (MSFlexGrid1.TextMatrix(contador, 8)) = "__/__/____" Then
        fecfab = Format(MSFlexGrid1.TextMatrix(contador, 9), "DD/MM/YYYY")
        uSql = "insert into STKLOTE (STSALMA,STSCODIGO,STSLOTE,STSLKDIS,STSFECFAB) VALUES ('" & alma & "','" & codigo & "','" & Lote & "'," & CANTIDAD & ",'" & fecfab & "') "
     ElseIf (MSFlexGrid1.TextMatrix(contador, 9)) = "__/__/____" And (MSFlexGrid1.TextMatrix(contador, 8)) <> "__/__/____" Then
        fecven = Format(MSFlexGrid1.TextMatrix(contador, 8), "DD/MM/YYYY")
        uSql = "insert into STKLOTE  (STSALMA,STSCODIGO,STSLOTE,STSLKDIS,STSFECVEN)  VALUES ('" & alma & "','" & codigo & "','" & Lote & "' ," & CANTIDAD & " ,'" & fecven & "') " 'SIN FECFAB
     ElseIf (MSFlexGrid1.TextMatrix(contador, 9)) <> "__/__/____" And (MSFlexGrid1.TextMatrix(contador, 8)) <> "__/__/____" Then
        'If Not IsDate(fecfab) Then
        fecfab = Date
        'If Not IsDate(fecven) Then
        fecven = Date
        
        uSql = "insert into STKLOTE  (STSALMA,STSCODIGO,STSLOTE,STSLKDIS,STSFECFAB,STSFECVEN)  VALUES ('" & alma & "','" & codigo & "','" & Lote & "' ," & CANTIDAD & " ,'" & fecfab & "','" & fecven & "') "
     Else
'        If Not IsDate(fecfab) Then
        fecfab = Date
'        If Not IsDate(fecven) Then
        fecven = Date
        
        uSql = "insert into STKLOTE  (STSALMA,STSCODIGO,STSLOTE,STSLKDIS)  VALUES ('" & alma & "','" & codigo & "','" & Lote & "' ," & CANTIDAD & ") "
     End If
    End If
    VGCNx.Execute uSql
       
End Sub

Private Sub grabaserie(alma As String, codigo As String)
Dim uSql As String
Dim Serie As String
Dim Valor As Integer
Dim rs As New ADODB.Recordset
Dim RSQL As String
Dim fecfab As Date
Dim fecven As Date
    Serie = MSFlexGrid1.TextMatrix(contador, 2)
    RSQL = "select STSSKDIS FROM STKSERI where STSALMA= '" & alma & "' and  STSCODIGO= '" & codigo & "' and STSSERIE= '" & Serie & "'" '
'    Set RS = VGBaseDatos.OpenRecordset(RSQL, dbOpenSnapshot)
    Set rs = VGCNx.Execute(RSQL)
    If rs.RecordCount > 0 Then
       Valor = IIf((Campo = "NI" And alma <> TxtAlmacen.text), 1, 0)
       uSql = "update STKSERI set STSSKDIS = " & Valor & " WHERE  STSALMA='" & alma & "' and STSCODIGO='" & codigo & "'AND STSSERIE='" & Serie & "'"
    Else
       If (Campo = "NI" And alma <> TxtAlmacen.text) Then
         uSql = "insert into STKSERI (STSALMA,STSCODIGO,STSSERIE,STSFECMOV,STSFECVEN,STSSKDIS,STSSKCOM) VALUES ('" & alma & "','" & codigo & "','" & Serie & "' ,'" & Date & "','" & Date & "',1,1) "
       Else
         uSql = "insert into STKSERI (STSALMA,STSCODIGO,STSSERIE,STSFECMOV,STSFECVEN,STSSKDIS,STSSKCOM) VALUES ('" & alma & "','" & codigo & "','" & Serie & "' ,'" & Date & "','" & Date & "',0,0) "
       End If
    End If
    VGCNx.Execute uSql
End Sub

Public Sub grabastk1(contador)
   Dim acmd As New ADODB.Command
   Dim criterio As String
   Dim cadena As String
   Dim ndato As Double
   Dim rsbusca As New ADODB.Recordset
   
   cadena = MSFlexGrid1.TextMatrix(contador, 0)
   criterio = " STCODIGO = '" & cadena & "'"
   criterio = criterio + "and  STALMA = '" & TxtAlmacen & "'"
 '  Data3.Recordset.FindFirst criterio
   Set rsbusca = VGCNx.Execute("SELECT * FROM STKART WHERE " & criterio)
   If rsbusca.RecordCount = 0 Then
'     Data3.Recordset.AddNew
'     Data3.Recordset("STSKDIS") = CANTIDAD
'     Data3.Recordset("STKPREPRO") = precioprom
'     Data3.Recordset("STALMA") = TxtAlmacen  '"01"
'     Data3.Recordset("STCODIGO") = MSFlexGrid1.TextMatrix(contador, 0)
      VGCNx.Execute "Insert Into Stkart " & _
                        "(STSKDIS,STKPREPRO,STALMA,STCODIGO)" & _
                        " values(" & _
                        CANTIDAD & "," & precioprom & ",'" & TxtAlmacen & "'," & MSFlexGrid1.TextMatrix(contador, 0) & "')"
                        
       'Grabamos en Facturacion
        Set acmd.ActiveConnection = VGCNx
        acmd.CommandText = "al_actualizaproducto_pro"
        acmd.CommandType = adCmdStoredProc
        acmd.Prepared = True
        With acmd
            .Parameters("@baseini") = VGCNx.DefaultDatabase
            .Parameters("@basefin") = VGBase2
            .Parameters("@almacen") = TxtAlmacen
            .Parameters("@articulo") = MSFlexGrid1.TextMatrix(contador, 0)
            .Parameters("@tipo") = "1"
        End With
        acmd.Execute
        Set acmd = Nothing
                        
   Else
     'Data3.Recordset.Edit
     auxdisp = rsbusca.Fields("STSKDIS")
     If rsbusca.Fields("STKPREPRO") <> 0 And (canttemp + auxdisp) <> 0 Then   'no se registrado algun precio
       'rsbusca.Fields("STKPREPRO") = VGTipCamb * (precioprom * canttemp + auxdisp * rsbusca.Fields("STKPREPRO")) / (canttemp + auxdisp)
       VGCNx.Execute "UPDATE stkart " & _
                          " SET STKPREPRO =" & VGTipCamb * (precioprom * canttemp + auxdisp * rsbusca.Fields("STKPREPRO")) / (canttemp + auxdisp) & _
                          " WHERE " & criterio
       
     End If
     'Data3.Recordset("STSKDIS") = Data3.Recordset("STSKDIS") + CANTIDAD
      VGCNx.Execute "UPDATE stkart " & _
                        " SET STSKDIS =" & rsbusca.Fields("STSKDIS") + CANTIDAD & _
                        " WHERE " & criterio
   End If
   rsbusca.Close
   Set rsbusca = Nothing
   
  ' Data3.Recordset.Update
   
   If MSFlexGrid1.TextMatrix(contador, 7) = "S" Then grabaserie TxtAlmacen, MSFlexGrid1.TextMatrix(contador, 0)
   If MSFlexGrid1.TextMatrix(contador, 7) = "N" Then grabalote TxtAlmacen, MSFlexGrid1.TextMatrix(contador, 0)
   'Data3.Refresh
   Call ValMes(TxtAlmacen, True)
End Sub

Private Sub devolver(NumDoc As String)
   Dim adors As New ADODB.Recordset
   Dim rs As New ADODB.Recordset
   Dim RSQL As String
   
   RSQL = "select  CACODCLI,CAFECDOC,CACODMON,CASITGUI  from MovAlmCab where CAALMA = '" & TxtAlmacen.text & "' and CATD= 'GS'  AND  CANUMDOC= '" & NumDoc & "' AND CASITGUI IN ( 'V','P')  AND NOT CACIERRE "
   Set adors = New ADODB.Recordset
   adors.Open RSQL, VGCNx, adOpenDynamic, adLockOptimistic
   If adors.RecordCount = 0 Then
       MsgBox "No existe el número, ha sido Anulado o ha sido Facturado o  se ha producido el Cierre Mensual", vbCritical, "Verificar"
       Exit Sub
   End If
   If VGGuiaSal = False Then
       Deshabilitar (False)
   Else
       Deshabilitar (True)
   End If
   DTPicker1 = adors(1)
   Text5 = adors("cacodcli")
   If Not IsNull(adors("cacodmon")) Then
      If adors("cacodmon") = "02" Then

      End If
   End If
   EstadoDevolucion = adors("casitgui")
   adors.Close
   RSQL = "select  CNOMCLI,CDIRCLI from MAECLI   where CCODCLI= '" & Trim(Text5.text) & "' "
   'Set RS = VGBaseDatos.OpenRecordset(RSQL, dbOpenSnapshot)
   Set rs = VGCNx.Execute(RSQL)
   If rs.RecordCount = 0 Then
       MsgBox "No existe Cliente, Documento Incompleto ", vbCritical, mensaje1
       Exit Sub
   End If
   Text6 = rs(0)
   Text7 = rs(1)
   Call llenarFG(Text3, Text4)
End Sub

Private Sub actualiza_guia_dev()
  Dim uSql As String
  uSql = "Update MovAlmCab set CASITGUI = 'A' where CAALMA = '" & TxtAlmacen.text & "' and CATD= 'GS'  AND  CANUMDOC= '" & Text4 & "'"
  VGCNx.Execute uSql
End Sub

Function buscarclie(doc As String) As Recordset
  Dim rs As New ADODB.Recordset
  Dim RSQL As String
  RSQL = "select  CNOMCLI,CDIRCLI from MAECLI   where CCODCLI= '" & doc & "' "
  'Set RS = VGBaseDatos.OpenRecordset(RSQL, dbOpenSnapshot)
  Set rs = VGCNx.Execute(RSQL)
  If rs.EOF Then
       MsgBox "No existe Cliente ", vbCritical, mensaje1
       Exit Function
  End If
End Function

Public Sub buscarstk(Cod As String, CANTIDAD As Double, suma As Boolean)
  Dim rs As New ADODB.Recordset
  Dim RSQL As String
  RSQL = "select n.STSKDIS from  StkArt  n.STALMA = '" & TxtAlmacen.text & "'   and n.STCODIGO= " & Cod & " "
  'Set RS = VGBaseDatos.OpenRecordset(RSQL, dbOpenSnapshot)
  Set rs = VGCNx.Execute(RSQL)
  If rs.EOF Then
     MsgBox "No hay dicho articulo en almacen", vbCritical, mensaje1
     Exit Sub
  End If
  If suma Then
     rs(0) = rs(0) + CANTIDAD
  Else
     rs(0) = rs(0) - CANTIDAD
  End If
End Sub

Private Sub llenarFG(tipo As String, NumDoc As String)
     Dim Adoreg1 As ADODB.Recordset
     Dim RSQL As String
     Dim ser_lot As String
     Dim dato As String
     ' FG.FormatString = "Codigo| Descripcion| Serie \ Lote|  Cantidad | A Devolver| A Entregar||"
      MSFlexGrid1.Row = 0
      MSFlexGrid1.Cols = 11
      MSFlexGrid1.ColWidth(0) = 1500
      MSFlexGrid1.ColWidth(1) = 2700
      MSFlexGrid1.ColWidth(2) = 1000
      MSFlexGrid1.ColWidth(3) = 1200
      MSFlexGrid1.ColWidth(4) = 1200
      MSFlexGrid1.ColWidth(5) = 1200
      MSFlexGrid1.ColWidth(6) = 2
      MSFlexGrid1.ColWidth(7) = 2
      MSFlexGrid1.ColWidth(8) = 1100
      MSFlexGrid1.ColWidth(9) = 1100
      MSFlexGrid1.ColWidth(10) = 1100
      
      MSFlexGrid1.ColAlignment(0) = 1
      RSQL = "select n.DECODIGO, n.DEDESCRI, m.AUNIDAD, n.DECANTID, n.DESERIE,n.DELOTE  from MovAlmDet n ,maeArt m where  n.DEALMA ='" & TxtAlmacen.text & "' AND n.DETD = '" & tipo & "' AND n.DENUMDOC ='" & NumDoc & "' and m.acodigo=n.decodigo  ORDER BY n.DEITEM "  '

     Set Adoreg1 = New ADODB.Recordset
     Adoreg1.Open RSQL, VGCNx, adOpenDynamic, adLockOptimistic
     If Adoreg1.RecordCount = 0 Then
       Exit Sub
     End If
     Adoreg1.MoveFirst
     MSFlexGrid1.Rows = 1
     While Not Adoreg1.EOF
       If IsNull(Adoreg1(4)) And IsNull(Adoreg1(5)) Then
             ser_lot = ""
             dato = "X"
       ElseIf Not IsNull(Adoreg1(4)) Then
             ser_lot = Adoreg1(4)
             dato = "S"
       Else
             ser_lot = Adoreg1(5)
             dato = "N"
       End If                '0               1               2                 3               4              5            6            7
       MSFlexGrid1.AddItem (Adoreg1(0) & vbTab & Adoreg1(1) & vbTab & ser_lot & vbTab & Adoreg1(3) & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & dato)
       Adoreg1.MoveNext
     Wend
End Sub
Private Sub visualizarFG1()
     
     MSFlexGrid1.Cols = 11
     MSFlexGrid1.Row = 0
     visualizarFG2
     MSFlexGrid1.ColWidth(0) = 1500  'cod
     MSFlexGrid1.ColWidth(1) = 2700   'des
     MSFlexGrid1.ColWidth(2) = 1200   'ser
     MSFlexGrid1.ColWidth(3) = 1500   'cant
     MSFlexGrid1.ColWidth(4) = 1300   'und
     MSFlexGrid1.ColWidth(5) = 1500   'vv.
     MSFlexGrid1.ColWidth(6) = 1500   'pv
     MSFlexGrid1.ColWidth(7) = 2   'xserie
     MSFlexGrid1.ColWidth(8) = 1100
     MSFlexGrid1.ColWidth(9) = 1100
     MSFlexGrid1.ColWidth(10) = 1100
     
     MSFlexGrid1.ColAlignment(0) = 1
End Sub
Private Sub visualizarFG2()
     'MsFlexGrid1.FormatString = "   Codigo |   Descripcion|  Serie\Lote| Cantidad|  Unidad | V.Valor |P.Valor|| "
     MSFlexGrid1.Cols = 12
     MSFlexGrid1.Row = 0
     MSFlexGrid1.TextMatrix(0, 0) = " CODIGO "
     MSFlexGrid1.TextMatrix(0, 1) = " DESCRIPCION"
     MSFlexGrid1.TextMatrix(0, 2) = " SERIE \ LOT"
     MSFlexGrid1.TextMatrix(0, 3) = " CANTIDAD "
     MSFlexGrid1.TextMatrix(0, 4) = " UNIDAD "
     MSFlexGrid1.TextMatrix(0, 5) = " V. VALOR"
     MSFlexGrid1.TextMatrix(0, 6) = " P. VALOR"
     MSFlexGrid1.TextMatrix(0, 7) = " F"   'revisar
     MSFlexGrid1.TextMatrix(0, 8) = "Cent.Costo "
     MSFlexGrid1.TextMatrix(0, 9) = "Ord.Fabri  "
     MSFlexGrid1.TextMatrix(0, 10) = "Maqu./Equi."
     MSFlexGrid1.TextMatrix(0, 11) = "Can.Ref"
     
     MSFlexGrid1.ColWidth(8) = 1100
     MSFlexGrid1.ColWidth(9) = 1100
     MSFlexGrid1.ColWidth(10) = 1100
     MSFlexGrid1.ColAlignment(0) = 1
End Sub

Private Sub visualizarFG()
  MSFlexGrid1.Clear
  MSFlexGrid1.Row = 0
  MSFlexGrid1.ColWidth(0) = 1500  'cod
  MSFlexGrid1.ColWidth(1) = 2700   'des
  MSFlexGrid1.ColWidth(2) = 1000   'ser
  MSFlexGrid1.ColWidth(3) = 1200   'cant
  MSFlexGrid1.ColWidth(4) = 1200   'und
  MSFlexGrid1.ColWidth(5) = 1200   'vv.
  MSFlexGrid1.ColWidth(6) = 2   'vv.
  MSFlexGrid1.ColWidth(8) = 1100
  MSFlexGrid1.ColWidth(9) = 1100
  MSFlexGrid1.ColWidth(10) = 1100
  
  MSFlexGrid1.Rows = 1
  MSFlexGrid1.TextMatrix(0, 0) = " CODIGO"
  MSFlexGrid1.TextMatrix(0, 1) = " DESCRIPCION"
  MSFlexGrid1.TextMatrix(0, 2) = " SERIE"
  MSFlexGrid1.TextMatrix(0, 3) = " CANTIDAD"
  MSFlexGrid1.TextMatrix(0, 4) = " A DEVOLVER"
  MSFlexGrid1.TextMatrix(0, 5) = " A ENTREGAR"
  MSFlexGrid1.TextMatrix(0, 8) = "Cent.Costo "
  MSFlexGrid1.TextMatrix(0, 9) = "Ord.Fabri  "
  MSFlexGrid1.TextMatrix(0, 10) = "Maqu./Equi."
  
  MSFlexGrid1.ColAlignment(0) = 1
End Sub

Private Sub llenadatos()
  'Text12 = ""
'  Label13 = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0)
'  Label15 = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1)
'  Label17 = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3)
'  Text12.SetFocus
End Sub

Private Sub reinicia()
  limpia
  MSFlexGrid1.Clear
  MSFlexGrid1.Rows = 1
  visualizarFG2
  Command1.Visible = False
  Command2.Visible = False
  Command3.Visible = False
'  CmdSalir.SetFocus
End Sub


Private Sub Deshabilitar(flag As Boolean)
  Text3.Enabled = flag
  Text4.Enabled = flag
  Text5.Enabled = flag
  Text6.Enabled = flag
  Text7.Enabled = flag
  Text8.Enabled = flag
  TxtAlmacen.Enabled = flag
End Sub

Function transa(text As TextBox) As String
 Dim rs As Recordset
 Dim RSQL As String
 Dim dato As String
  dato = "S"
  RSQL = "select  TT_DESCRI,tt_clie FROM TabTransa where TT_CODMOV= '" & text & "' and TT_TIPMOV ='S'"    '& dato & "'" '
  
   Set rs = VGCNx.Execute(RSQL)
  If rs.RecordCount > 0 Then
    transa = rs(0)
    Label9 = rs(0)
    Cliente = IIf(rs(1) = "S", True, False)
  Else
    MsgBox "El tipo de codigo no existe !", vbOKOnly + vbExclamation, "Error"
    transa = ""
  End If
   rs.Close
End Function

Function ValidarDoc(txt As TextBox) As String
  
  Dim rs As Recordset
  Dim RSQL As String
RSQL = "select TDO_DESCRI  from TIPO_DOCU  where TDO_TIPDOC='" & txt.text & "'"

Set rs = VGCNx.Execute(RSQL)
If rs.RecordCount = 0 Then
   MsgBox "El tipo de codigo no existe !", vbOKOnly, "Error"
   ValidarDoc = ""
   txt.SetFocus
   Exit Function
End If
ValidarDoc = rs(0)
Label10.Caption = rs(0)
rs.Close
End Function

Private Sub grabacabecera()
Dim uSql As String
Dim Data1 As New ADODB.Recordset
Dim empresaorigen As String
'If Text4.text <> "" Then
 On Error GoTo GrabErr
VGCNx.BeginTrans
Data1.Open "select * from tabalm where taalma='" & TxtAlmacen & "'", VGCNx, adOpenDynamic, adLockOptimistic
 
 If Data1.RecordCount > 0 Then
    numserie = Right("00000000000" & Trim(CStr(Data1!tanumsal)), 11)                 'nro pedido"
 End If
 Data1("tanumsal") = Data1("tanumsal") + 1
 empresaorigen = Data1!empresacodigo
Data1.UpdateBatch
Data1.Close
VGCNx.CommitTrans
Data1.Open "movalmcab", VGCNx, adOpenDynamic, adLockOptimistic
      Data1.AddNew
      Data1("CAALMA") = TxtAlmacen.text     '"0
      Data1("CATIPMOV") = "S"
      Data1("CATD") = "NS"
      Data1("CAUSUARI") = VGUsuario
      tipo = Data1("CATD")
      Data1("CACOTIZA") = IIf(Len(Trim(tx_ordfab)) = 0, " ", tx_ordfab)
      Data1("CARFTDOC") = Text3.text
      Data1("CARFNDOC") = nroguia
      Data1("CAFECDOC") = DTPicker1
      'guardar el nro del doc referencial
      Data1("CATIPGUI") = TxTransa
      Data1("CAHORA") = Format(Time, "hh:mm:ss")
      Data1("CAFECACT") = Now
      If Trim(TxTransa.text) <> "" Then
         Data1("CACODMOV") = UCase$(Trim(TxTransa.text))
      Else
         Data1("CACODMOV") = " "
      End If
      If Trim(Text6.text) <> "" Then
         Data1("CANOMCLI") = LTrim(Text6)
      Else
         Data1("CANOMCLI") = " "
      End If
      If Trim(Text7.text) <> "" Then
         Data1("CADIRENV") = LTrim(RTrim(Text7))
      Else
         Data1("CADIRENV") = " "
      End If
      If Text5.Visible And Trim(Text5.text) <> "" Then
         Data1("CACODCLI") = Trim(Text5.text)
         Data1("CARUC") = IIf(ruc <> "", ruc, " ")
      Else
         Data1("CACODCLI") = " "
         Data1("CARUC") = " "
      End If
      If Trim(Text8.text) <> "" Then
         Data1("CANUMORD") = Mid$(UCase$(Text8.text), 1, 10)
      Else
         Data1("CANUMORD") = " "
      End If
         Data1("CATIPCAM") = DevolverTCambio(DTPicker1.Value)
         Data1("CAFORVEN") = " "
      If Trim(TxtTransp) = "" Then
         Data1("CACODTRAN") = " "
      Else
          Data1("CACODTRAN") = Trim(TxtTransp)
      End If

      Data1("CAVENDE") = Ctr_AyuVendedor.xclave

      Data1("CARFALMA") = " "
      If Not VGGuiaSal Then                 'para devolucion
         Data1("CASITGUI") = Trim(EstadoDevolucion)
      Else
           Data1("CASITGUI") = "V"    'para guia de remision cualquiera
      End If
      Data1("CAESTIMP") = "V"
      Data1("CAFECACT") = Now
      Data1("CANUMDOC") = numserie
      Data1("canroped") = Txtnrodoc
      Data1("Contacto") = TxtCon
      Data1("caructra") = TextDNI.text
      Data1("canomtra") = TextRazon.text
      Data1("cadirtra") = TextDir.text
      Data1("empresacodigo") = empresaorigen
      Data1.Update
   'End If
   'Data1.Refresh
   Data1.Close
   Set Data1 = Nothing
   
   If Text3.text <> "GR" Then
      uSql = "Update NUM_DOCUMENTOS set   CTNNUMERO= " & numsal & " where  CTNCODIGO = 'GS' AND CTNNUMSER= '" & CmbSerie.text & "'"
  '   VGCNx.Execute uSql
   End If
   hubo_error = False
   Exit Sub
GrabErr:

    MsgBox Err.Description
    hubo_error = True
    Exit Sub
    Resume
End Sub

Private Sub buscar()
  Dim criterio As String
  Dim rs As Recordset
  
  Dim RSQL As String
  analitico = 0
   TxTransa = UCase(Trim(TxTransa))
   'Busco la transaccion
   RSQL = "select  *  from TabTransa  where TT_CODMOV ='" & TxTransa.text & "' and TT_TIPMOV ='S'"
   'Set RS = VGBaseDatos.OpenRecordset(RSQL, dbOpenSnapshot)
   Set rs = VGCNx.Execute(RSQL)
   If rs.RecordCount = 0 Then
      MsgBox "El tipo de transaccion no existe !", vbOKOnly, "Error"
      LIMPIACABECERA
      TxTransa.SetFocus
      Exit Sub
   End If

   If Not IsNull(rs("TT_CONT")) Then
            TT_CONTADOR = rs("TT_CONT")
   Else
       MsgBox "El tipo de transaccion no esta inicialida !", vbOKOnly, "Error"
       Exit Sub
   End If
 
   Label11.Visible = True
   TxtAlmacen.Visible = True
   If rs("TT_OC") = "N" Then
'      Text8.Enabled = False
   End If
   If rs("TT_CLIE") = "S" Then
         Text6.Visible = True
         Text7.Visible = True
         Text7.Enabled = True
   Else
         Text6.Visible = False
         Text7.Visible = False
         Text7.Enabled = False
         
   End If
   'MsgBox "Transaccion correcta", vbOKOnly, "Aviso"
   '*RMM*************************************
   If rs("TT_CC") = "N" Then
      Label27.Visible = True  'False
      FrmCreacionSal.txccosto.Visible = False
      FrmCreacionSal.lblccosto.Visible = False
   Else
      Label27.Visible = True
  
      FrmCreacionSal.txccosto.Visible = True
      FrmCreacionSal.lblccosto.Visible = True
   End If
   '*RMM*************************************
           
   If rs("TT_ORDFAB") = "S" Then
      tx_ordfab.Visible = True
      Label25.Visible = True
      FrmCreacionSal.lblordfab.Visible = True
      FrmCreacionSal.TxordFab.Visible = True
   Else
      tx_ordfab.Visible = False  'False
      Label25.Visible = False   'False
      FrmCreacionSal.lblordfab.Visible = False
      FrmCreacionSal.TxordFab.Visible = False
   End If
   
   If rs("TT_EQUIP") = "S" Then
      tx_codmaq.Visible = False
      Label26.Visible = False
      analitico = 1
      FrmCreacionSal.Ctr_AyuAnalitico.Enabled = True
      FrmCreacionSal.Ctr_AyuAnalitico.Visible = True
   Else
      tx_codmaq.Visible = False
      Label26.Visible = False    'False
      FrmCreacionSal.Ctr_AyuAnalitico.Enabled = False
      FrmCreacionSal.Ctr_AyuAnalitico.Visible = False
   End If
           
   'lbltrans = Mid(lbltrans, 1, 21)
   If Text3.Enabled Then
      Text3.SetFocus
   ElseIf Text5.Enabled Then
      Text5.SetFocus
   ElseIf Text6.Enabled Then
      Text6.SetFocus
   ElseIf Text7.Enabled Then
      Text7.SetFocus
   ElseIf Text8.Enabled Then
      Text8.SetFocus
   ElseIf TxtAlmacen.Enabled Then
      TxtAlmacen.SetFocus
   Else
      TxTransa.SetFocus
   End If
   CmdGrabarCab.Enabled = True
   
End Sub
Private Sub ValMes(almacen As String, entrada As Boolean)
  Dim cadena As String
  Dim criterio As String
 
  Dim adors As New ADODB.Recordset
  Dim RSQL As String
  Dim Cantent As Double
  Dim Cantsal As Double
  Dim mespro As String
  Dim uSql As String
   mespro = Year(DTPicker1) & Format(Month(DTPicker1), "00")
   cadena = MSFlexGrid1.TextMatrix(contador, 0) 'codigo del art
   RSQL = "select SMCANENT,SMCANSAL FROM MoResMes where  SMALMA = '" & almacen & "' and SMMESPRO = '" & mespro & "' AND  SMCODIGO= '" & cadena & "'"  '
   
  'Set adors = New ADODB.Recordset
  Set adors = VGCNx.Execute(RSQL)
  ' adors.Open RSQL, Vgcnx, adOpenDynamic, adLockOptimistic
  If adors.RecordCount <> 0 Then
      If Not VGGuiaSal Then
             Cantent = adors(1) - cantidadDEV   '1
             uSql = "Update MoResMes set SMCANSAL = " & Cantent & "  where SMALMA='" & almacen & "'  and  SMCODIGO ='" & cadena & "' AND SMMESPRO ='" & mespro & "' "
       Else
             If entrada Then
                    Cantent = adors(0) + CANTIDAD
                    uSql = "Update MoResMes set SMCANENT = " & Cantent & " where SMALMA='" & almacen & "' and  SMCODIGO ='" & cadena & "' AND SMMESPRO ='" & mespro & "' "
             Else
                    Cantsal = adors(1) + CANTIDAD
                    uSql = "Update MoResMes set SMCANSAL = " & Cantsal & " where SMALMA='" & almacen & "' and  SMCODIGO ='" & cadena & "' AND SMMESPRO ='" & mespro & "' "
             End If
       End If
   Else
       Cantsal = IIf(entrada, 0, CANTIDAD)
       Cantent = IIf(entrada, CANTIDAD, 0)
       uSql = "insert into MoResMes (SMALMA,SMCODIGO,SMMESPRO,SMCANENT,SMCANSAL,SMSALDOINI) VALUES ('" & almacen & "','" & cadena & "','" & mespro & "' ," & Cantent & "," & Cantsal & ",0) "
   End If
   VGCNx.Execute uSql
   adors.Close
End Sub


'******* ********************************  Factura  *************
Private Sub llenarfactura()
 Dim rs As ADODB.Recordset
 Dim RSQL As String
 Dim dato As String
 Dim NumDoc As String
 Dim numserie1 As String
 numserie1 = Mid(Text4, 1, 4)
 NumDoc = Mid(Text4, 4, 14)
 RSQL = "Select * from FACCAB where  cfnumser = '" & numserie1 & "' "  'and cfnrocaj ='" & vGPtoVenta & "'
 RSQL = RSQL & "and cfnumdoc = '" & NumDoc & "' AND cftd= '" & Text3 & "'"
 
 Set rs = New ADODB.Recordset
 rs.Open RSQL, VGCNx, adOpenStatic
 If rs.RecordCount > 0 Then
   If rs("CFFACGUI") = "S" Then  'cuando la graba
      MsgBox "Documento de referencia tiene guia, no procede", vbExclamation, "Aviso"
      rs.Close
      Exit Sub
   End If
   If Not IsNull(rs("CFFECDOC")) Then DTPicker1 = rs("CFFECDOC")
   If Not IsNull(rs("CFCODCLI")) Then Text5 = rs("CFCODCLI")
   If Not IsNull(rs("CFNOMBRE")) Then Text6 = rs("CFNOMBRE")
   If Not IsNull(rs("CFDIRECC")) Then Text7 = rs("CFDIRECC")
    If Not IsNull(rs("CFORDCOM")) Then Text8 = rs("CFORDCOM")
    detallefact
    CmdGrabarDet.Visible = True
Else
    MsgBox "No existe Factura", vbInformation, "Mensaje"
End If
rs.Close
End Sub
Private Sub detallefact()
 Dim rs As ADODB.Recordset
 Dim RSQL As String
 Dim dato As String
 Dim NumDoc As String
 Dim numserie1 As String
 Dim Serie As String
 numserie1 = Mid(Text4, 1, 3)  'obtengo la serie
 NumDoc = Mid(Text4, 4, 11)   'obtengo el numero de doc

 RSQL = "Select * from FACDET where  dfnumser = '" & numserie1 & "' "  'and cfnrocaj ='" & vGPtoVenta & "'                    '    A Inner Join FACCAB B on a.DFTD = B.CFTD and "
 RSQL = RSQL & "and dfnumdoc = '" & NumDoc & "'  AND    dftd ='" & Text3 & "' "
 Set rs = New ADODB.Recordset
 rs.Open RSQL, VGCNx, adOpenStatic
 If rs.RecordCount > 0 Then
   If MSFlexGrid1.Rows > 1 Then
        MSFlexGrid1.Rows = 1
   End If
   MSFlexGrid1.Refresh
   While Not rs.EOF
     If Not IsNull(rs("DFSERIE")) Then
             Serie = rs("DFSERIE")
     ElseIf Not IsNull(rs("DFLOTE")) Then
             Serie = rs("DFLOTE")
     Else
             Serie = ""
     End If
     MSFlexGrid1.AddItem (rs("dfcodigo") & vbTab & rs("dfdescri") & vbTab & Serie & vbTab & rs("dfcantid") & vbTab & rs("dfunidad") & vbTab & rs("dfprec_ven") & vbTab & rs("dfprec_ori"))
     rs.MoveNext
   Wend
 Else
   MsgBox "No existe el registro en Detalle de Factura", vbInformation, "Mensaje"
 End If
 rs.Close
End Sub

Function verifica_nro_guia(nroguia1 As String) As Boolean
Dim csql As String
Dim adors As ADODB.Recordset
   verifica_nro_guia = True
   If Len(nroguia1) <> 14 Then
     Exit Function
   End If

nroguia = CmbSerie.text & Text4
If Text3.text = "GR" Then

   SQL = "select puntovtadoccorr from vt_puntovtadocumento where empresacodigo='" & VGParametros.empresacodigo & "' and documentocodigo='GR' "
   SQL = SQL & " and  puntovtacodigo='" & VGParametros.puntovta & "' "
   SQL = SQL & " and puntovtadocserie='" & CmbSerie.text & "'"
     Set adors = VGCNx.Execute(SQL)
   If adors.RecordCount > 0 Then
      csql = Right("0000000000" & TraeDataSerie(SQL, VGCNx), 10)
      VGCNx.Execute "Update vt_puntovtadocumento " & _
       " set puntovtadoccorr='" & CStr(csql) + 1 & "'" & _
       " Where empresacodigo='" & VGParametros.empresacodigo & "' and documentocodigo='GR' and puntovtacodigo='" & VGParametros.puntovta & "' and puntovtadocserie='" & CmbSerie.text & "'"
      Text4 = csql
      nroguia = CmbSerie.text & csql
   End If
   csql = "SELECT CASITGUI FROM MovAlmCab where empresacodigo='" & VGParametros.empresacodigo & "' and CARFTDOC='GR'  and  CARFNDOC ='" & nroguia & "' "
   Set adors = New ADODB.Recordset
   adors.Open csql, VGCNx, adOpenDynamic, adLockOptimistic
   If adors.RecordCount = 0 Then
        verifica_nro_guia = False
   Else
        MsgBox "El número de Guia de remisión " & nroguia & "  ya fue grabada ", vbInformation, "Aviso"
        Exit Function
   End If
   
End If

   
End Function

Private Sub LIMPIACABECERA()
Label10.Caption = ""
Text3 = ""
Ctr_AyudaEmpresa.xclave = "": Ctr_AyudaEmpresa.xnombre = ""
Text5.text = ""
Text6.text = ""
Text7.text = ""
Text8.text = ""
Txtnrodoc.text = ""
TxtAlmacen.text = ""
TxtCon.text = ""
Texttipdoc.text = ""
Text1.text = ""
tx_ordfab.text = ""
ChkTalla.Value = False
Ctr_AyuVendedor.xclave = "": Ctr_AyuVendedor.xnombre = ""
Ctr_AyuTransporte.xclave = "": Ctr_AyuTransporte.xnombre = ""
tx_codmaq.text = ""
MSFlexGrid1.Rows = 1
End Sub


Private Sub GRABO_DET()
'(((((((((((((((((((((((((((
'esto es para facturacion
'Dim csql As String
'
'csql = "Insert Into FACWORDET (dfNUMPED,dfVENDE,dfNROCAJ,dfsecuen,dfCODIGO,"
'If Trim(TxArticulo) = "TEXTO" Then
'    csql = csql & "dfTEXTO,"
'Else
'    csql = csql & "dfDESCRI,"
'End If
'csql = csql & "DFPRECI1,DFPRECIO,DFPORDES,DFDESCTO,DFDESCLI,DFDESESP,DFCANTID,DFCANREF,DFPRESUP,DFORDEN,DFESTADO,"
'csql = csql & "DFFECDOC,DFIGV,DFIGVPOR,DFIMPUS,DFIMPMN,DFSERIE,DFLOTE,DFALMA) VALUES ("
'If cTip = "GS" Then
'    csql = csql & "'A',"
'Else
'    csql = csql & "'F',"
'End If
'csql = csql '" & FrmPFacD04.cNumPed & "','" & FrmPFacD04.TxVendedor & "','" & vGPtoVenta & "',"
'csql = csql & "'" & Format(Str(FrmPFacD04.nItem), "00") & "','" & TxArticulo & "','" & TxDescripcion & "',"
'csql = csql & "" & nPreci1 & "," & TxPrecio & "," & TxDescuento & "," & FrmPFacD04.nImpDsAr & "," & FrmPFacD04.nImpDsCl & ","
'csql = csql & "" & FrmPFacD04.nImpDsEs & "," & TxCantidad & "," & TxCantRefe & ",'" & TxPresupuesto & "','" & TxNum & "','V', '" & Format(FrmPFacD04.MaskEdBox1, "mm/dd/yyyy") & "',"
'csql = csql & "" & FrmPFacD04.nwIgv & "," & nIgvPor & "," & FrmPFacD04.nImpUS & "," & FrmPFacD04.nImpMN & ",'" & FrmPFacD04.TxSerie & "',"
'csql = csql & "'" & TxLote & "','" & FrmPFacD04.TxPto & "')"
'
'VGBaseDatos.Execute csql
'
End Sub

Private Sub crtlvisible(dato As Boolean)
   MSFlexGrid1.Visible = dato
   Command1.Visible = dato
   Command2.Visible = dato
   Command3.Visible = dato
   CmdGrabarDet.Visible = dato
   CmdSalir.Visible = dato
   
End Sub

Function existe_clie(text As TextBox) As String
  Dim RSQL As String
  Dim rs As New ADODB.Recordset
  direccion = ""
  'RSQL = "SELECT CNOMCLI ,CDIRCLI FROM maecli where CCODCLI= '" & text & "'"
  RSQL = "Select clienterazonsocial as cnomcli,clientedireccion as cdircli " & _
       " FROM vt_cliente where clientecodigo='" & text & "'"
  'Set RS = VGBaseDatos.OpenRecordset(RSQL, dbOpenSnapshot)
  Set rs = VGCNx.Execute(RSQL)
   If Not rs.EOF Then 'existe
     existe_clie = rs(0)
     direccion = IIf(IsNull(rs(1)), " ", rs(1))
   Else
     existe_clie = ""
  End If
  rs.Close
End Function


Private Sub TxtTransp_DblClick()
Dim Adodc3 As ADODB.Recordset
Set Adodc3 = New ADODB.Recordset
Adodc3.Open "SELECT TRACODIGO,TRANOMBRE FROM al_transporte", VGCNx, adOpenStatic, adLockOptimistic
frmReferencia.Conectar Adodc3, "SELECT TRACODIGO,TRANOMBRE FROM al_transporte"
frmReferencia.Label1.Caption = "Transportista"
frmReferencia.Show vbModal
Adodc3.Close
If vGUtil(1) <> "" Then TxtTransp = (vGUtil(1))

If TxtTransp <> "" Then TxtTransp_KeyPress (13)
End Sub

Private Sub TxtTransp_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
   TxtTransp_DblClick
ElseIf KeyCode = 48 Then
    
ElseIf KeyCode = 8 Then
    
End If
End Sub

Private Sub TxtTransp_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
         If TxtTransp <> "" Then
            Dim Adodc3 As ADODB.Recordset
            Set Adodc3 = New ADODB.Recordset
            TxtTransp = Trim(TxtTransp)
            Adodc3.Open "SELECT TRACODIGO,TRANOMBRE FROM al_transporte WHERE  TRACODIGO= '" & TxtTransp & "'", VGCNx, adOpenStatic, adLockOptimistic
            If Adodc3.EOF Then
              If vbYes = MsgBox("El código de Transportista no existe," & Chr(13) & "desea agregarlo ", vbInformation + vbYesNo, "Aviso") Then
                  VGtransp = False
                  FrmTranspor.Show 1
                  VGtransp = True
               Else
                 TxtTransp = ""
              End If
            Else
              
            End If
            If CmdGrabarCab.Visible = True And CmdGrabarCab.Enabled = True Then
              CmdGrabarCab.SetFocus
            End If
         Else
            SendKeys "{tab}"
         End If
 End If
End Sub

Private Sub AgregarSerie()
Dim Adodc3 As ADODB.Recordset
Set Adodc3 = New ADODB.Recordset
'TxTransa
Adodc3.Open "SELECT CTNNUMSER FROM NUM_DOCUMENTOS   WHERE CTNCODIGO = '" & "GR" & "'", VGCNx, adOpenStatic, adLockOptimistic
 While Not Adodc3.EOF
    CmbSerie.AddItem Adodc3(0)
    Adodc3.MoveNext
 Wend
 Adodc3.Close
 CmbSerie.Visible = True
 CmbSerie.Enabled = True
 'CmbSerie.ListIndex = 0
End Sub

Private Sub TxVendedor_DblClick()
Dim Adodc3 As ADODB.Recordset
Set Adodc3 = New ADODB.Recordset
Adodc3.Open "SELECT vendedorcodigo,vendedornombres  FROM vt_VENDEDOR", VGCNx, adOpenStatic, adLockOptimistic
frmReferencia.Conectar Adodc3, "SELECT vendedorcodigo,vendedornombres  FROM vt_VENDEDOR"
frmReferencia.Label1.Caption = "Vendedor"
frmReferencia.Show vbModal
Adodc3.Close

End Sub

Private Sub TxVendedor_GotFocus()
'Enfoque TxVendedor
End Sub

Private Sub TxVendedor_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then TxVendedor_DblClick
End Sub

Private Sub TxVendedor_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'    If Existe(1, TxVendedor, "vt_vendedor", "vendedorcodigo", False) Then
'        SendKeys "{tab}"
'    Else
'        If Trim(TxVendedor) = "" Then
'            MsgBox "Ingrese Vendedor", vbInformation, "Mensaje"
'        Else
'            MsgBox "El Vendedor no existe", vbInformation, "Mensaje"
'        End If
'        TxVendedor.SetFocus
'    End If
'End If
End Sub


