VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form FrmFormadePago 
   Caption         =   "Form1"
   ClientHeight    =   7050
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11250
   LinkTopic       =   "Form1"
   ScaleHeight     =   7050
   ScaleWidth      =   11250
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmbotones 
      Height          =   1050
      Left            =   1950
      TabIndex        =   44
      Top             =   5760
      Width           =   5655
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
         Picture         =   "FrmFormadePago.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   49
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
         Picture         =   "FrmFormadePago.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   48
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
         Picture         =   "FrmFormadePago.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   180
         Width           =   870
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
         Left            =   1320
         Picture         =   "FrmFormadePago.frx":0CC6
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   180
         Width           =   915
      End
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
         Picture         =   "FrmFormadePago.frx":1108
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   180
         Width           =   915
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5595
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10785
      _ExtentX        =   19024
      _ExtentY        =   9869
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Consulta"
      TabPicture(0)   =   "FrmFormadePago.frx":154A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "TDBGrid1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Mantenimiento"
      TabPicture(1)   =   "FrmFormadePago.frx":1566
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cCancela"
      Tab(1).Control(1)=   "cAcepta"
      Tab(1).Control(2)=   "Frame1"
      Tab(1).Control(3)=   "Frame2"
      Tab(1).ControlCount=   4
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
         Left            =   -70110
         TabIndex        =   42
         Top             =   4770
         Width           =   1335
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
         Left            =   -71790
         TabIndex        =   41
         Top             =   4770
         Width           =   1335
      End
      Begin VB.Frame Frame1 
         Height          =   1965
         Left            =   -74880
         TabIndex        =   20
         Top             =   360
         Width           =   10155
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
            Left            =   1890
            MaxLength       =   2
            TabIndex        =   30
            Top             =   210
            Width           =   615
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
            Left            =   4650
            MaxLength       =   50
            TabIndex        =   29
            Top             =   210
            Width           =   5355
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
            Left            =   4650
            MaxLength       =   30
            TabIndex        =   28
            Top             =   660
            Width           =   5355
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
            Height          =   360
            Index           =   3
            Left            =   1890
            MaxLength       =   1
            TabIndex        =   27
            Top             =   660
            Width           =   345
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
            Left            =   8640
            MaxLength       =   20
            TabIndex        =   26
            Top             =   1080
            Width           =   1245
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
            Height          =   360
            Index           =   7
            Left            =   4680
            MaxLength       =   20
            TabIndex        =   25
            Top             =   1080
            Width           =   1395
         End
         Begin VB.CheckBox chk 
            Height          =   195
            Index           =   9
            Left            =   1830
            TabIndex        =   24
            Top             =   1560
            Width           =   345
         End
         Begin VB.CheckBox chk 
            Height          =   195
            Index           =   10
            Left            =   4800
            TabIndex        =   23
            Top             =   1560
            Width           =   465
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
            Index           =   5
            Left            =   1890
            MaxLength       =   3
            TabIndex        =   22
            Top             =   1080
            Width           =   615
         End
         Begin VB.CheckBox chk 
            Height          =   195
            Index           =   7
            Left            =   8610
            TabIndex        =   21
            Top             =   1620
            Width           =   345
         End
         Begin VB.Label lbl 
            Caption         =   "Tipo Documento"
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
            Left            =   120
            TabIndex        =   40
            Top             =   210
            Width           =   1680
         End
         Begin VB.Label lbl 
            Caption         =   "Descripcion"
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
            Left            =   2640
            TabIndex        =   39
            Top             =   210
            Width           =   1200
         End
         Begin VB.Label lbl 
            Caption         =   "Descrip. Corta"
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
            Index           =   2
            Left            =   2640
            TabIndex        =   38
            Top             =   660
            Width           =   1440
         End
         Begin VB.Label lbl 
            Caption         =   "Cargo / Abono"
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
            Index           =   3
            Left            =   120
            TabIndex        =   37
            Top             =   660
            Width           =   1560
         End
         Begin VB.Label lbl 
            Caption         =   "Codigo Sunat"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   19
            Left            =   75
            TabIndex        =   36
            Top             =   1080
            Width           =   1410
         End
         Begin VB.Label lbl 
            Caption         =   "Cta Ctble. CC Dolares"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   15
            Left            =   6315
            TabIndex        =   35
            Top             =   1080
            Width           =   2250
         End
         Begin VB.Label lbl 
            Caption         =   "Cta Ctble. CC Soles"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   14
            Left            =   2595
            TabIndex        =   34
            Top             =   1080
            Width           =   2250
         End
         Begin VB.Label lbl 
            Caption         =   "Aplica Retencion"
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
            Index           =   18
            Left            =   2580
            TabIndex        =   33
            Top             =   1560
            Width           =   1740
         End
         Begin VB.Label lbl 
            Caption         =   "Nota Contable"
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
            Index           =   4
            Left            =   120
            TabIndex        =   32
            Top             =   1560
            Width           =   1620
         End
         Begin VB.Label lbl 
            Caption         =   "Aplica Dif.Cambio"
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
            Index           =   16
            Left            =   6360
            TabIndex        =   31
            Top             =   1560
            Width           =   1740
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Caracteristicas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   2175
         Left            =   -74880
         TabIndex        =   1
         Top             =   2400
         Width           =   10095
         Begin VB.CheckBox chk 
            Height          =   195
            Index           =   8
            Left            =   5490
            TabIndex        =   10
            Top             =   780
            Width           =   345
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
            Index           =   6
            Left            =   5490
            MaxLength       =   11
            TabIndex        =   9
            Top             =   1470
            Width           =   1365
         End
         Begin VB.CheckBox chk 
            Height          =   195
            Index           =   6
            Left            =   2850
            TabIndex        =   8
            Top             =   1560
            Width           =   345
         End
         Begin VB.CheckBox chk 
            Height          =   195
            Index           =   5
            Left            =   5490
            TabIndex        =   7
            Top             =   1170
            Width           =   345
         End
         Begin VB.CheckBox chk 
            Height          =   195
            Index           =   4
            Left            =   2850
            TabIndex        =   6
            Top             =   1200
            Width           =   345
         End
         Begin VB.CheckBox chk 
            Height          =   195
            Index           =   3
            Left            =   8850
            TabIndex        =   5
            Top             =   780
            Width           =   345
         End
         Begin VB.CheckBox chk 
            Height          =   315
            Index           =   2
            Left            =   2820
            TabIndex        =   4
            Top             =   780
            Width           =   285
         End
         Begin VB.CheckBox chk 
            Height          =   195
            Index           =   1
            Left            =   5490
            TabIndex        =   3
            Top             =   510
            Width           =   345
         End
         Begin VB.CheckBox chk 
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
            Left            =   2820
            TabIndex        =   2
            Top             =   390
            Width           =   255
         End
         Begin VB.Label lbl 
            Caption         =   "Permite Cancelacion"
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
            Index           =   17
            Left            =   3300
            TabIndex        =   19
            Top             =   780
            Width           =   2490
         End
         Begin VB.Label lbl 
            Caption         =   "Num. Correlativo"
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
            Index           =   13
            Left            =   3300
            TabIndex        =   18
            Top             =   1530
            Width           =   1950
         End
         Begin VB.Label lbl 
            Caption         =   "Num. Automatica"
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
            Index           =   12
            Left            =   180
            TabIndex        =   17
            Top             =   1560
            Width           =   2490
         End
         Begin VB.Label lbl 
            Caption         =   "Valida Banco"
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
            Index           =   11
            Left            =   3300
            TabIndex        =   16
            Top             =   1110
            Width           =   2490
         End
         Begin VB.Label lbl 
            Caption         =   "Docum. Renovac. Letras"
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
            Index           =   10
            Left            =   180
            TabIndex        =   15
            Top             =   1110
            Width           =   2490
         End
         Begin VB.Label lbl 
            Caption         =   "Permite Renovac. Letras"
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
            Left            =   6180
            TabIndex        =   14
            Top             =   780
            Width           =   2490
         End
         Begin VB.Label lbl 
            Caption         =   "Permite Aplicaciones"
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
            Index           =   8
            Left            =   180
            TabIndex        =   13
            Top             =   780
            Width           =   2250
         End
         Begin VB.Label lbl 
            Caption         =   "Ing. en Plan. Pagos"
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
            Index           =   6
            Left            =   3300
            TabIndex        =   12
            Top             =   420
            Width           =   2160
         End
         Begin VB.Label lbl 
            Caption         =   "Ing. en Plan. Apertura"
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
            Index           =   5
            Left            =   180
            TabIndex        =   11
            Top             =   360
            Width           =   2400
         End
      End
      Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
         Height          =   4935
         Left            =   240
         TabIndex        =   43
         Top             =   450
         Width           =   10305
         _ExtentX        =   18177
         _ExtentY        =   8705
         _LayoutType     =   0
         _RowHeight      =   15
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
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
         MultipleLines   =   0
         CellTipsWidth   =   104.882
         DeadAreaBackColor=   13160660
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
Attribute VB_Name = "FrmFormadePago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adll As New dllgeneral.dll_general
Dim modoinsert As Boolean
Dim modoedit As Boolean
Dim i_filaorigen As Integer
Dim nLongicampo(2) As Integer

Private Sub cAcepta_Click()
 If adll.VerificaDatoExistente(VGCNx, "select * from te_FormadePago Where FormadePagocodigo='" & txt(0) & "'") = 1 And modoinsert = True Then
    MsgBox "Ya existe el Codigo...!!!", vbInformation, MsgTitle
    Exit Sub
 End If

 If modoinsert = True Then
       VGCNx.Execute "Insert Into te_formadepago " & _
                  "(FormadePagocodigo,FormadePagodescripcion,FormadePagodesccorta, " & _
                  "FormadePagotipo,FormadePagoingplan,FormadePagoingcobra,FormadePagopermiteaplica," & _
                  "FormadePagorenovarletras,FormadePagodocrenovaletra,FormadePagovalidabanco," & _
                  "FormadePagonumeauto,FormadePagonumerador,FormadePagocuentasoles,FormadePagocuentadolares," & _
                  "FormadePagoaplicadifcamb,FormadePagonotaconta,documentoretencion," & _
                  "FormadePagosunat,usuariocodigo,fechaact,FormadePagocancela)" & _
                  "VALUES(" & _
                  "'" & txt(0) & "','" & txt(1) & "'," & _
                  "'" & txt(2) & "','" & txt(3) & "'," & _
                  "'" & IIf(chk(0).Value = 1, "1", "0") & "'," & _
                  "'" & IIf(chk(1).Value = 1, "1", "0") & "'," & _
                  "'" & IIf(chk(2).Value = 1, "1", "0") & "'," & _
                  "'" & IIf(chk(3).Value = 1, "1", "0") & "'," & _
                  "'" & IIf(chk(4).Value = 1, "1", "0") & "'," & _
                  "'" & IIf(chk(5).Value = 1, "1", "0") & "'," & _
                  "'" & IIf(chk(6).Value = 1, "1", "0") & "'," & _
                  "'" & txt(6) & "','" & txt(7) & "','" & txt(8) & "'," & _
                  "'" & IIf(chk(7).Value = 1, "1", "0") & "'," & _
                  "'" & IIf(chk(9).Value = 1, "1", "0") & "'," & _
                  "'" & IIf(chk(10).Value = 1, "1", "0") & "'," & _
                  "'" & txt(5) & "','" & VGusuario & "','" & Format(Date, "dd/mm/yyyy") & "'," & _
                  "'" & IIf(chk(8).Value = 1, "1", "0") & "')"
 
 ElseIf modoedit = True Then
       VGCNx.Execute "Update te_FormadePago " & _
                  " Set  FormadePagodescripcion='" & txt(1) & "'," & _
                  "FormadePagodesccorta='" & txt(2) & "'," & _
                  "FormadePagotipo='" & txt(3) & "'," & _
                  "FormadePagoingplan='" & IIf(chk(0).Value = 1, "1", "0") & "'," & _
                  "FormadePagoingcobra='" & IIf(chk(1).Value = 1, "1", "0") & "'," & _
                  "FormadePagopermiteaplica='" & IIf(chk(2).Value = 1, "1", "0") & "'," & _
                  "FormadePagorenovarletras='" & IIf(chk(3).Value = 1, "1", "0") & "'," & _
                  "FormadePagodocrenovaletra='" & IIf(chk(4).Value = 1, "1", "0") & "'," & _
                  "FormadePagovalidabanco='" & IIf(chk(5).Value = 1, "1", "0") & "'," & _
                  "FormadePagonumeauto='" & IIf(chk(6).Value = 1, "1", "0") & "'," & _
                  "FormadePagonumerador='" & txt(6) & "'," & _
                  "FormadePagocuentasoles='" & txt(7) & "'," & _
                  "FormadePagocuentadolares='" & txt(8) & "'," & _
                  "FormadePagoaplicadifcamb='" & IIf(chk(7).Value = 1, "1", "0") & "'," & _
                  "FormadePagonotaconta='" & IIf(chk(9).Value = 1, "1", "0") & "'," & _
                  "documentoretencion='" & IIf(chk(10).Value = 1, "1", "0") & "'," & _
                  "FormadePagosunat='" & txt(5) & "'," & _
                  "usuariocodigo='" & VGusuario & "'," & _
                  "fechaact='" & Format(Date, "dd/mm/yyyy") & "'," & _
                  "FormadePagocancela='" & IIf(chk(8).Value = 1, "1", "0") & "' " & _
                  " Where FormadePagocodigo='" & txt(0) & "'"
 
 End If
 modoedit = False
 modoinsert = False
 Call Listado
End Sub

Public Function Listado()
    TDBGrid1.ClearFields
    Set TDBGrid1.DataSource = Nothing
    Call adll.ListarEnTDBGRID(VGCNx, "te_FormadePago", TDBGrid1, "FormadePagocodigo,FormadePagodescripcion,FormadePagorenovarletras as Renovacion_Letras,FormadePagodocrenovaletra as Doc_Renova,FormadePagovalidabanco,FormadePagonumeauto", "FormadePagocodigo", nLongicampo)
    Call ConfiguraGrid
    Call adll.ActivaTab(0, 1, SSTab1)
    frmbotones.Visible = True

End Function

Private Sub cCancela_Click()
  Call adll.ActivaTab(0, 1, SSTab1)
  frmbotones.Visible = True
End Sub

Private Sub cmdBotones_Click(Index As Integer)
  Dim J As Integer
  Dim spos As Integer
  Dim SQL As String
  Dim d_estado As Double
  ''''''''''
  Dim rs As New ADODB.Recordset
  Dim error As ADODB.Errors
  '''''''''''
  On Error GoTo CONTROLERRORES
  
  SSTab1.TabEnabled(1) = True

 modoedit = False
 modoinsert = False

  Select Case Index
  
     Case 0   'nuevo
        SSTab1.Tab = 1
        If txt(0).Visible = True Then
            txt(0).SetFocus
        ElseIf chk(0).Visible = True Then
            chk(0).SetFocus
        End If
        Call Limpia_textos
        Call adll.ActivaTab(1, 1, SSTab1)
        
        frmbotones.Visible = False
        modoinsert = True
        txt(0).SetFocus
        
     Case 1   'modificar
        If TDBGrid1.Row < 0 Then
          Exit Sub
        End If
        
        Call Limpia_textos
        
        Set rs = VGCNx.Execute("select * from te_FormadePago Where FormadePagocodigo='" & TDBGrid1.Columns(0).Text & "'")
        If rs.RecordCount > 0 Then
           txt(0) = Escadena(rs!FormadePagocodigo)
           txt(1) = Escadena(rs!FormadePagodescripcion)
           txt(2) = Escadena(rs!FormadePagodesccorta)
           txt(3) = Escadena(rs!FormadePagotipo)
           chk(0).Value = IIf(Escadena(rs!FormadePagoingplan) = "1", 1, 0)
           chk(1).Value = IIf(Escadena(rs!FormadePagoingcobra) = "1", 1, 0)
           chk(2).Value = IIf(Escadena(rs!FormadePagopermiteaplica) = "1", 1, 0)
           chk(3).Value = IIf(Escadena(rs!FormadePagorenovarletras) = "1", 1, 0)
           chk(4).Value = IIf(Escadena(rs!FormadePagodocrenovaletra) = "1", 1, 0)
           chk(5).Value = IIf(Escadena(rs!FormadePagovalidabanco) = "1", 1, 0)
           chk(6).Value = IIf(Escadena(rs!FormadePagonumeauto) = "1", 1, 0)
           chk(8).Value = IIf(Escadena(rs!FormadePagocancela) = "1", 1, 0)
           txt(6) = Escadena(rs!FormadePagonumerador)
           txt(7) = Escadena(rs!FormadePagocuentasoles)
           txt(8) = Escadena(rs!FormadePagocuentadolares)
           chk(7).Value = IIf(Escadena(rs!FormadePagoaplicadifcamb) = "1", 1, 0)
           chk(9).Value = IIf(Escadena(rs!FormadePagonotaconta) = "1", 1, 0)
           chk(10).Value = IIf(Escadena(rs!documentoretencion) = "1", 1, 0)
           txt(5) = Escadena(rs!FormadePagosunat)
        End If
        rs.Close
        Set rs = Nothing
        Call adll.ActivaTab(1, 1, SSTab1)
        frmbotones.Visible = False
        SSTab1.Tab = 1
        
        i_filaorigen = TDBGrid1.Row
        modoedit = True
        If txt(0).Visible = True Then
            txt(0).SetFocus
        ElseIf chk(0).Visible = True Then
            chk(0).SetFocus
        End If

     Case 2   'eliminar
          If TDBGrid1.Row < 0 Then
              Exit Sub
          End If
         
          If MsgBox("Desea Eliminar el Registro?", vbYesNo, MsgTitle) = vbYes Then
              VGCNx.Execute "Delete From  te_FormadePago where FormadePagocodigo='" & TDBGrid1.Columns(0).Text & "'"
          End If
          Call Listado
     Case 3  'Imprimir
       Call Imprimir("cp_RepMantTipoDocumento.rpt")
     Case 4  ' salir
       Unload Me
  End Select
  
  
'RaiseEvent Click(Index)

Exit Sub

CONTROLERRORES:
   If Err Then
       MsgBox Err.Number & "-" & Err.Description, vbInformation, "ERROR"
       Err = 0
       'VGgeneral.RollbackTrans
       Resume Next
    End If

End Sub

Public Function Limpia_textos()
 Dim J As Integer
   For J = 0 To txt.Count - 1
      If J <> 4 Then txt(J).Text = ""
   Next J
   For J = 0 To chk.Count - 1
      chk(J).Value = 0
   Next J
End Function

Private Sub Form_Load()
   MostrarForm Me, "C2"
   Call adll.ActivaTab(0, 1, SSTab1)
   nLongicampo(1) = 0
   Call adll.ListarEnTDBGRID(VGCNx, "te_FormadePago", TDBGrid1, "FormadePagocodigo,FormadePagodescripcion,FormadePagorenovarletras as Renovacion_Letras,FormadePagodocrenovaletra as Doc_Renova,FormadePagovalidabanco,FormadePagonumeauto", "FormadePagocodigo", nLongicampo)
   Call ConfiguraGrid
   
End Sub

Public Function ConfiguraGrid()
   With TDBGrid1
    .Columns(0).Width = 1200
    .Columns(0).Caption = "Codigo"
    .Columns(1).Width = 2500
    .Columns(1).Caption = "Descripcion"
    .Columns(2).Width = 1000
    .Columns(2).Caption = "Ren. Letras"
    .Columns(3).Width = 1000
    .Columns(3).Caption = "Doc.Renovar"
    .Columns(4).Width = 1000
    .Columns(4).Caption = "V. Banco"
    .Columns(5).Width = 1400
    .Columns(5).Caption = "Numer.Autom."
    .Refresh
   End With
End Function

Private Sub txt_Change(Index As Integer)
  Select Case Index
   Case 0, 6, 7, 8, 5, 4
      If Not adll.ValidaCadena(txt(Index), "N") Then
        If Len(Trim$(txt(Index))) > 0 Then
          txt(Index) = Left$(txt(Index), Len(txt(Index)) - 1)
        End If
        txt(Index).SetFocus
      End If
      Exit Sub
   Case 1, 2
      If Not adll.ValidaCadena(txt(Index), "C") Then
        If Len(Trim$(txt(Index))) > 0 Then
          txt(Index) = Left$(txt(Index), Len(txt(Index)) - 1)
        End If
        txt(Index).SetFocus
      End If
      Exit Sub
   Case 3
     ' If adll.ValidaCadena(txt(Index), "C") Then
         If Not UCase$(txt(Index)) Like "[AC]" Then
            If Len(Trim$(txt(Index))) > 0 Then
              txt(Index) = Left$(txt(Index), Len(txt(Index)) - 1)
            End If
            txt(Index).SetFocus
         End If
      'End If
      Exit Sub
  
  End Select

End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
 
 If KeyAscii = 13 Then
   txt(Index) = UCase$(txt(Index))
   If Index = 3 Then
     
     If Not UCase$(txt(Index)) Like "[AC]" Then
         If Len(Trim$(txt(Index))) > 0 Then
           txt(Index) = Left$(txt(Index), Len(txt(Index)) - 1)
         End If
         txt(Index).SetFocus
         Exit Sub
      End If
   ElseIf Index = 4 Then
     chk(0).SetFocus
     Exit Sub
   End If
   Call Seguir(txt(Index), KeyAscii)
 End If
 
End Sub

Private Sub txt_LostFocus(Index As Integer)
   If Index = 6 Then
       txt(Index) = Right$("000000000000000" & txt(Index), txt(Index).MaxLength)
   ElseIf Index = 0 Then
       txt(Index) = Right$("000000000000000" & txt(Index), txt(Index).MaxLength)
   ElseIf Index Like "[123]" Then
       txt(Index) = UCase$(txt(Index))
   End If
End Sub


