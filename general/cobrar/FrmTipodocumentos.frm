VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form FrmTipodocumentos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tipos de Documentos"
   ClientHeight    =   9045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9045
   ScaleWidth      =   9495
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   7755
      Left            =   210
      TabIndex        =   0
      Top             =   0
      Width           =   9105
      _ExtentX        =   16060
      _ExtentY        =   13679
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
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
      TabPicture(0)   =   "FrmTipodocumentos.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "TDBGrid1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Mantenimiento"
      TabPicture(1)   =   "FrmTipodocumentos.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "cCancela"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cAcepta"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame1"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      Begin VB.Frame Frame1 
         Height          =   6525
         Left            =   210
         TabIndex        =   8
         Top             =   390
         Width           =   8595
         Begin VB.CheckBox chk 
            Height          =   195
            Index           =   9
            Left            =   6570
            TabIndex        =   46
            Top             =   1770
            Width           =   345
         End
         Begin VB.CheckBox chk 
            Height          =   195
            Index           =   8
            Left            =   3030
            TabIndex        =   45
            Top             =   4260
            Width           =   345
         End
         Begin VB.CheckBox chk 
            Height          =   195
            Index           =   7
            Left            =   6720
            TabIndex        =   43
            Top             =   6060
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
            Left            =   3000
            MaxLength       =   20
            TabIndex        =   23
            Top             =   5880
            Width           =   1725
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
            Left            =   3000
            MaxLength       =   20
            TabIndex        =   22
            Top             =   5400
            Width           =   1755
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
            Left            =   6390
            MaxLength       =   14
            TabIndex        =   21
            Top             =   4950
            Width           =   1965
         End
         Begin VB.CheckBox chk 
            Height          =   195
            Index           =   6
            Left            =   3030
            TabIndex        =   20
            Top             =   5040
            Width           =   345
         End
         Begin VB.CheckBox chk 
            Height          =   195
            Index           =   5
            Left            =   3030
            TabIndex        =   19
            Top             =   4650
            Width           =   345
         End
         Begin VB.CheckBox chk 
            Height          =   195
            Index           =   4
            Left            =   3030
            TabIndex        =   18
            Top             =   3840
            Width           =   345
         End
         Begin VB.CheckBox chk 
            Height          =   195
            Index           =   3
            Left            =   3030
            TabIndex        =   17
            Top             =   3420
            Width           =   345
         End
         Begin VB.CheckBox chk 
            Height          =   315
            Index           =   2
            Left            =   3000
            TabIndex        =   16
            Top             =   2910
            Width           =   645
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
            Left            =   6600
            MaxLength       =   2
            TabIndex        =   15
            Top             =   2520
            Width           =   615
         End
         Begin VB.CheckBox chk 
            Height          =   195
            Index           =   1
            Left            =   3000
            TabIndex        =   14
            Top             =   2550
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
            Left            =   3000
            TabIndex        =   13
            Top             =   2070
            Width           =   255
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
            Left            =   2970
            MaxLength       =   1
            TabIndex        =   12
            Top             =   1590
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
            Index           =   2
            Left            =   2970
            MaxLength       =   30
            TabIndex        =   11
            Top             =   1140
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
            Index           =   1
            Left            =   2970
            MaxLength       =   50
            TabIndex        =   10
            Top             =   690
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
            Index           =   0
            Left            =   2940
            MaxLength       =   2
            TabIndex        =   9
            Top             =   210
            Width           =   615
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
            Left            =   360
            TabIndex        =   44
            Top             =   4170
            Width           =   2490
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
            Left            =   4830
            TabIndex        =   42
            Top             =   6000
            Width           =   1740
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
            Left            =   360
            TabIndex        =   41
            Top             =   5970
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
            Height          =   285
            Index           =   14
            Left            =   360
            TabIndex        =   40
            Top             =   5430
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
            Left            =   4440
            TabIndex        =   39
            Top             =   5010
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
            Left            =   360
            TabIndex        =   38
            Top             =   5010
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
            Left            =   360
            TabIndex        =   37
            Top             =   4590
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
            Left            =   360
            TabIndex        =   36
            Top             =   3750
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
            Left            =   360
            TabIndex        =   35
            Top             =   3330
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
            Left            =   360
            TabIndex        =   34
            Top             =   2910
            Width           =   2250
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
            Height          =   285
            Index           =   7
            Left            =   4770
            TabIndex        =   33
            Top             =   2580
            Width           =   1650
         End
         Begin VB.Label lbl 
            Caption         =   "Ing. en Plan. Cobr."
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
            Left            =   360
            TabIndex        =   32
            Top             =   2460
            Width           =   2400
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
            Left            =   360
            TabIndex        =   31
            Top             =   2040
            Width           =   2400
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
            Left            =   4740
            TabIndex        =   30
            Top             =   1710
            Width           =   1620
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
            Left            =   360
            TabIndex        =   29
            Top             =   1620
            Width           =   2400
         End
         Begin VB.Label lbl 
            Caption         =   "Descripcion Corta"
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
            Left            =   360
            TabIndex        =   28
            Top             =   1200
            Width           =   2400
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
            Left            =   360
            TabIndex        =   26
            Top             =   720
            Width           =   2400
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
            Left            =   360
            TabIndex        =   24
            Top             =   270
            Width           =   2400
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
         Left            =   3210
         TabIndex        =   25
         Top             =   7050
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
         Left            =   4890
         TabIndex        =   27
         Top             =   7050
         Width           =   1335
      End
      Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
         Height          =   6975
         Left            =   -74760
         TabIndex        =   1
         Top             =   450
         Width           =   8625
         _ExtentX        =   15214
         _ExtentY        =   12303
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
   Begin VB.Frame frmbotones 
      Height          =   1050
      Left            =   2070
      TabIndex        =   2
      Top             =   7800
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
         Picture         =   "FrmTipodocumentos.frx":0038
         Style           =   1  'Graphical
         TabIndex        =   7
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
         Left            =   1320
         Picture         =   "FrmTipodocumentos.frx":047A
         Style           =   1  'Graphical
         TabIndex        =   6
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
         Picture         =   "FrmTipodocumentos.frx":08BC
         Style           =   1  'Graphical
         TabIndex        =   5
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
         Picture         =   "FrmTipodocumentos.frx":0CFE
         Style           =   1  'Graphical
         TabIndex        =   4
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
         Picture         =   "FrmTipodocumentos.frx":1140
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   180
         Width           =   870
      End
   End
End
Attribute VB_Name = "FrmTipodocumentos"
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
 If adll.VerificaDatoExistente(VGCNx, "select * from cc_tipodocumento Where tdocumentocodigo='" & txt(0) & "'") = 1 And modoinsert = True Then
    MsgBox "Ya existe el Codigo...!!!", vbInformation, MsgTitle
    Exit Sub
 End If

 If modoinsert = True Then
       VGCNx.Execute "Insert Into cc_tipodocumento " & _
                  "(tdocumentocodigo,tdocumentodescripcion,tdocumentodesccorta, " & _
                  "tdocumentotipo,tdocumentoingplan,tdocumentoingcobra,tdocumentopermiteaplica," & _
                  "tdocumentorenovarletras,tdocumentodocrenovaletra,tdocumentovalidabanco," & _
                  "tdocumentonumeauto,tdocumentonumerador,tdocumentocuentasoles,tdocumentocuentadolares," & _
                  "tdocumentoaplicadifcamb,tdocumentonotaconta,tdocumentosunat,usuariocodigo,fechaact,tdocumentocancela)" & _
                  "VALUES(" & _
                  "'" & txt(0) & "'," & _
                  "'" & txt(1) & "'," & _
                  "'" & txt(2) & "'," & _
                  "'" & txt(3) & "'," & _
                  "'" & IIf(chk(0).Value = 1, "1", "0") & "'," & _
                  "'" & IIf(chk(1).Value = 1, "1", "0") & "'," & _
                  "'" & IIf(chk(2).Value = 1, "1", "0") & "'," & _
                  "'" & IIf(chk(3).Value = 1, "1", "0") & "'," & _
                  "'" & IIf(chk(4).Value = 1, "1", "0") & "'," & _
                  "'" & IIf(chk(5).Value = 1, "1", "0") & "'," & _
                  "'" & IIf(chk(6).Value = 1, "1", "0") & "'," & _
                  "'" & txt(6) & "'," & _
                  "'" & txt(7) & "'," & _
                  "'" & txt(8) & "'," & _
                  "'" & IIf(chk(7).Value = 1, "1", "0") & "'," & _
                  "'" & IIf(chk(9).Value = 1, "1", "0") & "'," & _
                  "'" & txt(5) & "','" & g_usuario & "','" & Format(Date, "dd/mm/yyyy") & "'," & _
                  "'" & IIf(chk(8).Value = 1, "1", "0") & "')"
 
 ElseIf modoedit = True Then
       VGCNx.Execute "Update cc_tipodocumento " & _
                  " Set  tdocumentodescripcion='" & txt(1) & "'," & _
                  "tdocumentodesccorta='" & txt(2) & "'," & _
                  "tdocumentotipo='" & txt(3) & "'," & _
                  "tdocumentoingplan='" & IIf(chk(0).Value = 1, "1", "0") & "'," & _
                  "tdocumentoingcobra='" & IIf(chk(1).Value = 1, "1", "0") & "'," & _
                  "tdocumentopermiteaplica='" & IIf(chk(2).Value = 1, "1", "0") & "'," & _
                  "tdocumentorenovarletras='" & IIf(chk(3).Value = 1, "1", "0") & "'," & _
                  "tdocumentodocrenovaletra='" & IIf(chk(4).Value = 1, "1", "0") & "'," & _
                  "tdocumentovalidabanco='" & IIf(chk(5).Value = 1, "1", "0") & "'," & _
                  "tdocumentonumeauto='" & IIf(chk(6).Value = 1, "1", "0") & "'," & _
                  "tdocumentonumerador='" & txt(6) & "'," & _
                  "tdocumentocuentasoles='" & txt(7) & "'," & _
                  "tdocumentocuentadolares='" & txt(8) & "'," & _
                  "tdocumentoaplicadifcamb='" & IIf(chk(7).Value = 1, "1", "0") & "'," & _
                  "tdocumentonotaconta='" & IIf(chk(9).Value = 1, "1", "0") & "'," & _
                  "tdocumentosunat='" & txt(5) & "'," & _
                  "usuariocodigo='" & g_usuario & "'," & _
                  "fechaact='" & Format(Date, "dd/mm/yyyy") & "'," & _
                  "tdocumentocancela='" & IIf(chk(8).Value = 1, "1", "0") & "' " & _
                  " Where tdocumentocodigo='" & txt(0) & "'"
 
 End If
 modoedit = False
 modoinsert = False
 Call Listado
End Sub


'FIXIT: Declare 'Listado' con un tipo de datos de enlace en tiempo de compilación          FixIT90210ae-R1672-R1B8ZE
Public Function Listado()
    TDBGrid1.ClearFields
    Set TDBGrid1.DataSource = Nothing
    Call adll.ListarEnTDBGRID(VGCNx, "cc_tipodocumento", TDBGrid1, "tdocumentocodigo,tdocumentodescripcion,tdocumentorenovarletras as Renovacion_Letras,tdocumentodocrenovaletra as Doc_Renova,tdocumentovalidabanco,tdocumentonumeauto", "tdocumentocodigo", nLongicampo)
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
        
        Set rs = VGCNx.Execute("select * from cc_tipodocumento Where tdocumentocodigo='" & TDBGrid1.Columns(0).Text & "'")
        If rs.RecordCount > 0 Then
           txt(0) = Escadena(rs!tdocumentocodigo)
           txt(1) = Escadena(rs!tdocumentodescripcion)
           txt(2) = Escadena(rs!tdocumentodesccorta)
           txt(3) = Escadena(rs!tdocumentotipo)
           chk(0).Value = IIf(Escadena(rs!tdocumentoingplan) = "1", 1, 0)
           chk(1).Value = IIf(Escadena(rs!tdocumentoingcobra) = "1", 1, 0)
           chk(2).Value = IIf(Escadena(rs!tdocumentopermiteaplica) = "1", 1, 0)
           chk(3).Value = IIf(Escadena(rs!tdocumentorenovarletras) = "1", 1, 0)
           chk(4).Value = IIf(Escadena(rs!tdocumentodocrenovaletra) = "1", 1, 0)
           chk(5).Value = IIf(Escadena(rs!tdocumentovalidabanco) = "1", 1, 0)
           chk(6).Value = IIf(Escadena(rs!tdocumentonumeauto) = "1", 1, 0)
           chk(8).Value = IIf(Escadena(rs!tdocumentocancela) = "1", 1, 0)
           txt(6) = Escadena(rs!tdocumentonumerador)
           txt(7) = Escadena(rs!tdocumentocuentasoles)
           txt(8) = Escadena(rs!tdocumentocuentadolares)
           chk(7).Value = IIf(Escadena(rs!tdocumentoaplicadifcamb) = "1", 1, 0)
           chk(9).Value = IIf(Escadena(rs!tdocumentonotaconta) = "1", 1, 0)
           txt(5) = Escadena(rs!tdocumentosunat)
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
              VGCNx.Execute "Delete From  cc_tipodocumento where tdocumentocodigo='" & TDBGrid1.Columns(0).Text & "'"
          End If
          Call Listado
     Case 3  'Imprimir
       Call Imprimir("RepMantTipoDocumento.rpt")
     Case 4  ' salir
       Unload Me
  End Select
  
  
'RaiseEvent Click(Index)

Exit Sub

CONTROLERRORES:
   If Err Then
       MsgBox Err.Number & "-" & Err.Description, vbInformation, "ERROR"
       Err = 0
       'VGGeneral.RollbackTrans
       Resume Next
    End If

End Sub


'FIXIT: Declare 'Limpia_textos' con un tipo de datos de enlace en tiempo de compilación     FixIT90210ae-R1672-R1B8ZE
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
   Call adll.ListarEnTDBGRID(VGCNx, "cc_tipodocumento", TDBGrid1, "tdocumentocodigo,tdocumentodescripcion,tdocumentorenovarletras as Renovacion_Letras,tdocumentodocrenovaletra as Doc_Renova,tdocumentovalidabanco,tdocumentonumeauto", "tdocumentocodigo", nLongicampo)
   Call ConfiguraGrid
   
End Sub



'FIXIT: Declare 'ConfiguraGrid' con un tipo de datos de enlace en tiempo de compilación     FixIT90210ae-R1672-R1B8ZE
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
        If Len(RTrim$(txt(Index))) > 0 Then
          txt(Index) = Left$(txt(Index), Len(txt(Index)) - 1)
        End If
        txt(Index).SetFocus
      End If
      Exit Sub
   Case 1, 2
      If Not adll.ValidaCadena(txt(Index), "C") Then
        If Len(RTrim$(txt(Index))) > 0 Then
          txt(Index) = Left$(txt(Index), Len(txt(Index)) - 1)
        End If
        txt(Index).SetFocus
      End If
      Exit Sub
   Case 3
     ' If adll.ValidaCadena(txt(Index), "C") Then
'FIXIT: Reemplazar la función 'UCase' con la función 'UCase$'.                             FixIT90210ae-R9757-R1B8ZE
         If Not UCase$(txt(Index)) Like "[AC]" Then
            If Len(RTrim$(txt(Index))) > 0 Then
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
'FIXIT: Reemplazar la función 'UCase' con la función 'UCase$'.                             FixIT90210ae-R9757-R1B8ZE
   txt(Index) = UCase$(txt(Index))
   If Index = 3 Then
     
'FIXIT: Reemplazar la función 'UCase' con la función 'UCase$'.                             FixIT90210ae-R9757-R1B8ZE
     If Not UCase$(txt(Index)) Like "[AC]" Then
         If Len(RTrim$(txt(Index))) > 0 Then
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
'FIXIT: Reemplazar la función 'UCase' con la función 'UCase$'.                             FixIT90210ae-R9757-R1B8ZE
       txt(Index) = UCase$(txt(Index))
   End If
End Sub
