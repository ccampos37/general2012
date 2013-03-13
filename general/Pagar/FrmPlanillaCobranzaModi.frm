VERSION 5.00
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form FrmPlanillaCobranzaModi 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Eliminar Documentos de Planilla de Cobranza"
   ClientHeight    =   8640
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11955
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   11955
   StartUpPosition =   3  'Windows Default
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
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Datos Generales"
      TabPicture(0)   =   "FrmPlanillaCobranzaModi.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(1)=   "Frame4"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Detalle"
      TabPicture(1)   =   "FrmPlanillaCobranzaModi.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "frmbotones"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame4 
         Height          =   930
         Left            =   -70284
         TabIndex        =   0
         Top             =   5400
         Width           =   1980
         Begin VB.CommandButton cmdBotones 
            Caption         =   "&Acepta"
            Height          =   690
            Index           =   11
            Left            =   96
            Picture         =   "FrmPlanillaCobranzaModi.frx":0038
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
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   180
            Width           =   855
         End
      End
      Begin VB.Frame frmbotones 
         Height          =   930
         Left            =   5220
         TabIndex        =   46
         Top             =   6660
         Width           =   2100
         Begin VB.CommandButton cmdBotones 
            Caption         =   "&Eliminar"
            Height          =   690
            Index           =   2
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   44
            Top             =   180
            Width           =   825
         End
         Begin VB.CommandButton cmdBotones 
            Caption         =   "&Salir"
            Height          =   690
            Index           =   4
            Left            =   1080
            Style           =   1  'Graphical
            TabIndex        =   45
            Top             =   180
            Width           =   870
         End
      End
      Begin VB.Frame Frame3 
         Height          =   6045
         Left            =   210
         TabIndex        =   11
         Top             =   570
         Width           =   11175
         Begin VB.Frame Frame5 
            Height          =   1335
            Left            =   150
            TabIndex        =   12
            Top             =   4620
            Width           =   10875
            Begin VB.Frame Frame6 
               Height          =   555
               Left            =   60
               TabIndex        =   27
               Top             =   720
               Width           =   10755
               Begin VB.Label Label3 
                  BackColor       =   &H00800000&
                  BorderStyle     =   1  'Fixed Single
                  ForeColor       =   &H00C0FFC0&
                  Height          =   285
                  Index           =   1
                  Left            =   180
                  TabIndex        =   29
                  Top             =   180
                  Width           =   10425
               End
               Begin VB.Label Label3 
                  Caption         =   "Saldo Doc"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000040C0&
                  Height          =   135
                  Index           =   0
                  Left            =   6090
                  TabIndex        =   30
                  Top             =   240
                  Width           =   975
               End
               Begin VB.Label Label3 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00800000&
                  BorderStyle     =   1  'Fixed Single
                  ForeColor       =   &H000080FF&
                  Height          =   225
                  Index           =   2
                  Left            =   6270
                  TabIndex        =   28
                  Top             =   210
                  Width           =   1305
               End
            End
            Begin VB.TextBox Text1 
               Height          =   285
               Index           =   6
               Left            =   5220
               MaxLength       =   4
               TabIndex        =   26
               Top             =   450
               Width           =   585
            End
            Begin VB.TextBox Text1 
               Height          =   285
               Index           =   5
               Left            =   4560
               MaxLength       =   2
               TabIndex        =   25
               Top             =   450
               Width           =   405
            End
            Begin VB.TextBox Text1 
               Height          =   285
               Index           =   4
               Left            =   4110
               MaxLength       =   1
               TabIndex        =   24
               Top             =   450
               Width           =   345
            End
            Begin VB.TextBox Text1 
               Height          =   285
               Index           =   3
               Left            =   2700
               MaxLength       =   10
               TabIndex        =   23
               Top             =   450
               Width           =   1155
            End
            Begin VB.TextBox Text1 
               Height          =   285
               Index           =   2
               Left            =   2100
               MaxLength       =   4
               TabIndex        =   22
               Top             =   450
               Width           =   525
            End
            Begin VB.TextBox Text1 
               Height          =   285
               Index           =   1
               Left            =   1530
               MaxLength       =   2
               TabIndex        =   21
               Top             =   450
               Width           =   465
            End
            Begin VB.TextBox Text1 
               Height          =   285
               Index           =   0
               Left            =   180
               MaxLength       =   11
               TabIndex        =   20
               Top             =   450
               Width           =   1275
            End
            Begin VB.TextBox Text1 
               Height          =   285
               Index           =   7
               Left            =   5910
               MaxLength       =   10
               TabIndex        =   19
               Top             =   450
               Width           =   1125
            End
            Begin VB.TextBox Text1 
               Height          =   285
               Index           =   8
               Left            =   7170
               MaxLength       =   2
               TabIndex        =   18
               Top             =   450
               Width           =   435
            End
            Begin VB.TextBox Text1 
               Height          =   285
               Index           =   9
               Left            =   7830
               MaxLength       =   12
               TabIndex        =   17
               Top             =   450
               Width           =   435
            End
            Begin VB.TextBox Text1 
               Height          =   285
               Index           =   10
               Left            =   8490
               MaxLength       =   8
               TabIndex        =   16
               Top             =   450
               Width           =   1065
            End
            Begin VB.TextBox Text1 
               Height          =   285
               Index           =   11
               Left            =   9630
               MaxLength       =   8
               TabIndex        =   15
               Top             =   450
               Width           =   1065
            End
            Begin VB.CommandButton cAyuda 
               Caption         =   "..."
               Height          =   315
               Index           =   0
               Left            =   3870
               TabIndex        =   14
               Top             =   450
               Width           =   195
            End
            Begin VB.CommandButton cAyuda 
               Caption         =   "..."
               Height          =   315
               Index           =   1
               Left            =   4980
               TabIndex        =   13
               Top             =   450
               Width           =   195
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
               Left            =   7800
               TabIndex        =   42
               Top             =   180
               Width           =   585
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
               Left            =   8490
               TabIndex        =   41
               Top             =   180
               Width           =   1005
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
               Left            =   7170
               TabIndex        =   40
               Top             =   180
               Width           =   465
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
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
               Index           =   5
               Left            =   4740
               TabIndex        =   39
               Top             =   180
               Width           =   285
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               Caption         =   "P/T"
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
               Left            =   4140
               TabIndex        =   38
               Top             =   180
               Width           =   315
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
               Left            =   2730
               TabIndex        =   37
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
               Left            =   2100
               TabIndex        =   36
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
               Left            =   1620
               TabIndex        =   35
               Top             =   180
               Width           =   315
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               Caption         =   "Proveedor"
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
               TabIndex        =   34
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
               Height          =   240
               Index           =   9
               Left            =   5940
               TabIndex        =   33
               Top             =   180
               Width           =   1125
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
               Height          =   240
               Index           =   10
               Left            =   5250
               TabIndex        =   32
               Top             =   180
               Width           =   615
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               Caption         =   "T. Cambio"
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
               Index           =   11
               Left            =   9750
               TabIndex        =   31
               Top             =   210
               Width           =   945
            End
         End
         Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
            Height          =   4305
            Left            =   150
            TabIndex        =   43
            Top             =   300
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
         BackColor       =   &H00C0C0C0&
         Height          =   3015
         Left            =   -72810
         TabIndex        =   6
         Top             =   2250
         Width           =   7695
         Begin VB.Frame Frame1 
            BackColor       =   &H00FF8080&
            BorderStyle     =   0  'None
            Height          =   2790
            Left            =   -408
            TabIndex        =   7
            Top             =   150
            Width           =   7545
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
               Left            =   3285
               TabIndex        =   1
               Top             =   1140
               Width           =   1185
               _ExtentX        =   2090
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   10
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuda2 
               Height          =   315
               Left            =   3285
               TabIndex        =   47
               Top             =   1710
               Width           =   3645
               _ExtentX        =   6429
               _ExtentY        =   556
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
            Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuda1 
               Height          =   285
               Left            =   3285
               TabIndex        =   48
               Top             =   600
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
               Top             =   660
               Width           =   2085
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "FECHA DE COBRANZA"
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
               Top             =   1200
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
               ForeColor       =   &H00C0FFC0&
               Height          =   225
               Index           =   2
               Left            =   960
               TabIndex        =   8
               Top             =   1740
               Width           =   2085
            End
         End
      End
   End
   Begin MSComctlLib.StatusBar Panel 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   2
      Top             =   8295
      Width           =   11955
      _ExtentX        =   21087
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
Attribute VB_Name = "FrmPlanillaCobranzaModi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adll As New dllgeneral.dll_general
Dim rsdetacmodi As New ADODB.Recordset




Private Sub cmdBotones_Click(Index As Integer)
  Dim acmd As New ADODB.Command
  Dim rb As New ADODB.Recordset
  Dim xabono, xzona, xmone, xcuenta As String
  Dim xnumplan, ximpsol, xtcam, xnumpag As Double
  On Error GoTo nerror
    Select Case Index
    Case 0
       Limpiartexto Text1, 0, 11
       Text1(0).SetFocus
       
    Case 2   'ELIMINAR DATOS
        If rsdetacmodi.RecordCount > 0 Then
            If MsgBox("Desea Eliminar el Registro?", vbYesNo, "AVISO") = vbNo Then
               Exit Sub
            End If
            ximpsol = CDbl(TDBGrid1.Columns(10).Text)
            xtcam = TDBGrid1.Columns(11).Text
            If TDBGrid1.Columns(8).Text <> xmone Then
               If TDBGrid1.Columns(8).Text = g_TipoSol Then
                  xtcam = TDBGrid1.Columns(11).Text
                  If TDBGrid1.Columns(11).Text = 0 Then xtcam = 1
                  ximpsol = CDbl(TDBGrid1.Columns(10).Text) / CDbl(xtcam)
               Else
                  xtcam = TDBGrid1.Columns(11).Text
                  If TDBGrid1.Columns(11).Text = 0 Then xtcam = 1
                   ximpsol = CDbl(TDBGrid1.Columns(10).Text) * CDbl(xtcam)
               End If
            End If
                            
                            
            DoEvents
                            
            '**** Actualizamos Saldos de documento pendiente
            If TDBGrid1.Columns(8).Text = g_TipoDolar Then
               If xmone = g_TipoSol Then
                   VGCNx.Execute "Update  cp_cargo Set cargoapeimppag=isnull(cargoapeimppag,0)-" & CDbl(TDBGrid1.Columns(10).Text / xtcam) & _
                              " Where documentocargo='" & TDBGrid1.Columns(1).Text & "' and cargonumdoc='" & Trim$(TDBGrid1.Columns(2).Text & TDBGrid1.Columns(3).Text) & "'"
               Else
                   VGCNx.Execute "Update  cp_cargo Set cargoapeimppag=isnull(cargoapeimppag,0)-" & CDbl(TDBGrid1.Columns(10).Text) & _
                              " Where documentocargo='" & TDBGrid1.Columns(1).Text & "' and cargonumdoc='" & Trim$(TDBGrid1.Columns(2).Text & TDBGrid1.Columns(3).Text) & "'"
               End If
            ElseIf TDBGrid1.Columns(8).Text = g_TipoSol Then
               If xmone = g_TipoDolar Then
                   VGCNx.Execute "Update  cp_cargo Set cargoapeimppag=isnull(cargoapeimppag,0)-" & CDbl(TDBGrid1.Columns(10).Text * xtcam) & _
                              " Where documentocargo='" & TDBGrid1.Columns(1).Text & "' and cargonumdoc='" & Trim$(TDBGrid1.Columns(2).Text & TDBGrid1.Columns(3).Text) & "'"
               Else
                   VGCNx.Execute "Update  cp_cargo Set cargoapeimppag=isnull(cargoapeimppag,0)-" & CDbl(TDBGrid1.Columns(10).Text) & _
                              " Where documentocargo='" & TDBGrid1.Columns(1).Text & "' and cargonumdoc='" & Trim$(TDBGrid1.Columns(2).Text & TDBGrid1.Columns(3).Text) & "'"
               End If
            End If
            
            VGCNx.Execute "Update  cp_cargo " & _
                        " Set cargoapeflgcan= CASE isnull(cargoapeimpape,0)-isnull(cargoapeimppag,0) WHEN 0 THEN '1' ELSE '0' END ," & _
                        "   cargoapefeccan='" & Date & "'" & _
                        " Where documentocargo='" & TDBGrid1.Columns(1).Text & "' and cargonumdoc='" & Trim$(TDBGrid1.Columns(2).Text & TDBGrid1.Columns(3).Text) & "'"
            
            
            'Eliminamos abono del documento
            
             Set rb = VGCNx.Execute("Select * from cp_abono where " & _
                        "documentoabono='" & TDBGrid1.Columns(1).Text & "' and " & _
                        "abononumdoc='" & Trim$(TDBGrid1.Columns(2).Text & TDBGrid1.Columns(3).Text) & "' and " & _
                        "abonotipoplanilla='" & Escadena(Ctr_Ayuda1.xclave) & "' and " & _
                        "abonocancli='" & TDBGrid1.Columns(0).Text & "' and  " & _
                        "abonocanfecpla='" & MBox1.Text & "' and " & _
                        "abonocantdqc='" & TDBGrid1.Columns(5).Text & "' and " & _
                        "abonocanndqc='" & Trim$(TDBGrid1.Columns(6).Text & TDBGrid1.Columns(7).Text) & "'")
                        
             If rb.RecordCount > 0 Then
                 xcuenta = Escadena(rb!abononumplanilla)
             End If
             rb.Close
             Set rb = Nothing
            
             VGCNx.Execute "insert into sysseguridad  values ('" & Date & "','" & Time & "','" & VGusuario & "','" & _
                       " Registro Eliminado... Documento : " & TDBGrid1.Columns(1).Text & "-" & Trim$(TDBGrid1.Columns(2).Text & TDBGrid1.Columns(3).Text) & _
                       " Planilla  : " & Escadena(Ctr_Ayuda1.xclave) & "- " & xcuenta & _
                       " Cliente   : " & TDBGrid1.Columns(0).Text & _
                       " Fecha     : " & MBox1.Text & _
                       " D.Cancela : " & TDBGrid1.Columns(5).Text & "-" & Trim$(TDBGrid1.Columns(6).Text & TDBGrid1.Columns(7).Text) & _
                       " Moneda    : " & TDBGrid1.Columns(8).Text & _
                       " Importe   : " & Numero(TDBGrid1.Columns(10).Text) & "')"
            

            VGCNx.Execute "delete from cp_abono where " & _
                       "documentoabono='" & TDBGrid1.Columns(1).Text & "' and " & _
                       "abononumdoc='" & Trim$(TDBGrid1.Columns(2).Text & TDBGrid1.Columns(3).Text) & "' and " & _
                       "abonotipoplanilla='" & Escadena(Ctr_Ayuda1.xclave) & "' and " & _
                       "abonocancli='" & TDBGrid1.Columns(0).Text & "' and  " & _
                       "abonocanfecpla='" & MBox1.Text & "' and " & _
                       "abonocantdqc='" & TDBGrid1.Columns(5).Text & "' and " & _
                       "abonocanndqc='" & Trim$(TDBGrid1.Columns(6).Text & TDBGrid1.Columns(7).Text) & "'"
            
            If rsdetacmodi.RecordCount >= 0 Then
              TDBGrid1.Delete
              TDBGrid1.Update
              TDBGrid1.Refresh
            End If
       End If
       
       MsgBox "El registro ha sido eliminado satisfactoriamente...!!!", vbInformation, MsgTitle
               
    Case 4
       Call adll.ActivaTab(0, 1, SSTab1)
    Case 11
      If Len(Trim$(Ctr_Ayuda1.xclave)) = 0 Then
        MsgBox "Falta Ingresar Tipo de Planilla...Verifique!!", vbInformation, MsgTitle
        Exit Sub
      End If
      If Len(Trim$(Ctr_Ayuda2.xclave)) = 0 Then
        MsgBox "Falta Ingresar Oficina/Vendedor...Verifique!!", vbInformation, MsgTitle
        Exit Sub
      End If
      If adll.VerificaDatoExistente(VGCNx, "select * from cp_tipoplanilla where tplanillacobranza='1' and tplanillacodigo='" & Escadena(Ctr_Ayuda1.xclave) & "' ") = 0 Then
            MsgBox "La planilla no es valida para realizar la cobranza...Verifique!!!", vbInformation, MsgTitle
            Ctr_Ayuda1.SetFocus
            Exit Sub
      End If

      If adll.VerificaDatoExistente(VGCNx, "select * from cp_abono where abonotipoplanilla='" & Escadena(Ctr_Ayuda1.xclave) & "' and vendedorcodigo='" & Ctr_Ayuda2.xclave & "' and abonocanfecpla='" & MBox1 & "' ") = 0 Then
         MsgBox "No existe planilla de esa fecha y/o vendedor...Verifique!!!", vbInformation, MsgTitle
         Exit Sub
      End If
      nAyuda = "1"
      If Date <> MBox1.Text Then
      '   Frmseguridad.Show 1
      Else
         nAyuda = "1"
      End If
      If nAyuda = "1" Then
        Call cargar_abonos("select * from cp_abono where abonotipoplanilla='" & Escadena(Ctr_Ayuda1.xclave) & "' and vendedorcodigo='" & Ctr_Ayuda2.xclave & "' and abonocanfecpla='" & MBox1 & "' ")
              
        Limpiartexto Text1, 0, 11
        Call adll.ActivaTab(1, 1, SSTab1)
        'Frame5.Enabled = True
      End If
      nAyuda = ""
    Case 12, 4
      Unload Me
  End Select
  
nerror:
  If Err Then
    MsgBox "Error : " & Err.Number & "-" & Err.Description, vbInformation, MsgTitle
    Err = 0
    Exit Sub
  End If
End Sub

Public Function cargar_abonos(nsql)
     Dim rb As New ADODB.Recordset
     
      Set rsdetacmodi = Nothing
      TDBGrid1.ClearFields
      Set TDBGrid1.DataSource = Nothing
      Call cargar_grilla
       
      Set rb = VGCNx.Execute(nsql)
      If rb.RecordCount > 0 Then
          rb.MoveFirst
          Do Until rb.EOF
            rsdetacmodi.AddNew
            rsdetacmodi.Fields(0) = rb!abonocancli
            rsdetacmodi.Fields(1) = rb!documentoabono
            rsdetacmodi.Fields(2) = Left$(rb!abononumdoc, 4)
            rsdetacmodi.Fields(3) = Right$(RTrim$(rb!abononumdoc), 10)
            rsdetacmodi.Fields(4) = rb!abonocanforcan
            rsdetacmodi.Fields(5) = rb!abonocantdqc
            rsdetacmodi.Fields(6) = Left$(rb!abonocanndqc, 4)
            rsdetacmodi.Fields(7) = Right$(RTrim$(rb!abonocanndqc), 10)
            rsdetacmodi.Fields(8) = rb!abonocanmoneda
            rsdetacmodi.Fields(9) = rb!abonocanbco
            rsdetacmodi.Fields(10) = Numero(rb!abonocanimpcan)
            rsdetacmodi.Fields(11) = Numero(rb!abonocantipcam)
            rsdetacmodi.Update
            rb.MoveNext
          Loop
      End If
      rb.Close
      Set rb = Nothing

End Function

Private Sub Form_Load()
  MostrarForm Me, "C"
  
  MBox1 = Format(Date, "DD/MM/YYYY")
  Call Ctr_Ayuda1.conexion(VGCNx)
  Call Ctr_Ayuda2.conexion(VGCNx)
  Ctr_Ayuda1.Filtro = "tplanillacobranza='1'"
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
        .Columns(11).Width = 800
        .Columns(11).NumberFormat = "###,###,##0.00"
        .Refresh
    End With
End Sub

Public Sub cargar_grilla()
   Set rsdetacmodi = Nothing
   Call rsdetacmodi.Fields.Append("Cliente", adChar, 11)
   Call rsdetacmodi.Fields.Append("TD", adChar, 2)
   Call rsdetacmodi.Fields.Append("Serie", adChar, 4)
   Call rsdetacmodi.Fields.Append("Numero", adChar, 10)
   Call rsdetacmodi.Fields.Append("P/T", adChar, 1)
   Call rsdetacmodi.Fields.Append("TDp", adChar, 2)
   Call rsdetacmodi.Fields.Append("Seriep", adChar, 4)
   Call rsdetacmodi.Fields.Append("Numerop", adChar, 10)
   Call rsdetacmodi.Fields.Append("Moneda", adChar, 2)
   Call rsdetacmodi.Fields.Append("Banco", adChar, 2)
   Call rsdetacmodi.Fields.Append("Importe", adDouble)
   Call rsdetacmodi.Fields.Append("TCambio", adDouble)
   
   rsdetacmodi.Open
   Set TDBGrid1.DataSource = rsdetacmodi
   TDBGrid1.Refresh
   Call ConfigGrid
End Sub


Private Sub Form_Unload(Cancel As Integer)
'  rsdetacmodi.Close
  
'  Set rsdetacmodi = Nothing
End Sub

Private Sub MBox1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then
     SendKeys "{tab}"
  End If
  
End Sub

Private Sub TDBGrid1_HeadClick(ByVal ColIndex As Integer)
   Dim nsql As String
   If TDBGrid1.ApproxCount >= 0 Then
   
     nsql = "select abonocancli,documentoabono,abononumdoc,abononumdoc,abonocanforcan,abonocantdqc,abonocanndqc,abonocanndqc,abonocanmoneda,abonocanbco,abonocanimpcan,abonocantipcam " & _
            " from cp_abono where abonotipoplanilla='" & Escadena(Ctr_Ayuda1.xclave) & "' and vendedorcodigo='" & Ctr_Ayuda2.xclave & "' and abonocanfecpla='" & MBox1 & "' order by " & CStr(ColIndex + 1)
     Call cargar_abonos(nsql)
   End If
 
End Sub

Private Sub TDBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
  Dim J As Integer
  Dim rb As New ADODB.Recordset
  Dim xpago, xcam As Double
 
 Frame5.Enabled = False
 If rsdetacmodi.RecordCount > 0 Then
    For J = 0 To rsdetacmodi.Fields.Count - 1
       Text1(J) = Escadena(TDBGrid1.Columns(J).Text)
    Next J
     
    Set rb = VGCNx.Execute("select * from cp_proveedor where clientecodigo='" & Trim$(Escadena(Text1(0))) & "'")
    If rb.RecordCount > 0 Then
       Label3(1) = Escadena(rb!clientecodigo) & "-" & Escadena(rb!clienterazonsocial)
    End If
    rb.Close
      
    Set rb = VGCNx.Execute("select * from cp_tipodocumento where tdocumentocodigo='" & Text1(1) & "' and tdocumentotipo='C'")
    If rb.RecordCount > 0 Then
'        MsgBox "El documento no es valido....Verifique!!!", vbInformation, MsgTitle
        rb.Close
        Set rb = Nothing
        Exit Sub
    End If
    rb.Close
    
    Text1(2) = Right$("000000000" & Trim$(Text1(2)), Text1(2).MaxLength)
    Text1(6) = Right$("000000000" & Trim$(Text1(6)), Text1(6).MaxLength)
    
 End If
   
 Set rb = Nothing
   
End Sub

Private Sub Text1_Change(Index As Integer)
  If Index = 0 Then
    If Len(Trim$(Text1(0))) = 0 Then
       Label3(1) = ""
    End If
    Label3(2) = ""
  End If
End Sub

Private Sub Text1_GotFocus(Index As Integer)
  Call adll.Enfoquetexto(Text1(Index))
End Sub

