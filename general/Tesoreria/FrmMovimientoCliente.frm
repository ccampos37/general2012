VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form FrmMovimientoClientes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Control de Cuentas por Cobrar"
   ClientHeight    =   9450
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9450
   ScaleWidth      =   10455
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   8010
      Left            =   210
      TabIndex        =   12
      Top             =   120
      Width           =   9975
      Begin VB.Frame Frame3 
         Height          =   5250
         Left            =   180
         TabIndex        =   50
         Top             =   2670
         Width           =   9675
         Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
            Height          =   3405
            Left            =   90
            TabIndex        =   51
            Top             =   210
            Width           =   9465
            _ExtentX        =   16695
            _ExtentY        =   6006
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
            AllowUpdate     =   0   'False
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
            _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=236,.bold=0,.fontsize=825,.italic=0"
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
         Begin VB.Frame Frame4 
            Height          =   1635
            Left            =   90
            TabIndex        =   52
            Top             =   3570
            Width           =   9465
            Begin VB.TextBox Text2 
               BackColor       =   &H80000000&
               Enabled         =   0   'False
               Height          =   285
               Index           =   0
               Left            =   180
               MaxLength       =   2
               TabIndex        =   53
               Top             =   390
               Width           =   465
            End
            Begin VB.TextBox Text2 
               Height          =   285
               Index           =   1
               Left            =   780
               MaxLength       =   2
               TabIndex        =   18
               Top             =   390
               Width           =   465
            End
            Begin VB.TextBox Text2 
               Height          =   285
               Index           =   2
               Left            =   1560
               MaxLength       =   14
               TabIndex        =   20
               Top             =   390
               Width           =   1125
            End
            Begin VB.TextBox Text2 
               Height          =   285
               Index           =   3
               Left            =   3060
               MaxLength       =   1
               TabIndex        =   22
               Top             =   390
               Width           =   285
            End
            Begin VB.TextBox Text2 
               Height          =   285
               Index           =   4
               Left            =   3480
               MaxLength       =   2
               TabIndex        =   23
               Top             =   390
               Width           =   435
            End
            Begin VB.TextBox Text2 
               Enabled         =   0   'False
               Height          =   285
               Index           =   5
               Left            =   4260
               MaxLength       =   2
               TabIndex        =   25
               Top             =   390
               Width           =   465
            End
            Begin VB.TextBox Text2 
               Enabled         =   0   'False
               Height          =   285
               Index           =   6
               Left            =   5040
               MaxLength       =   11
               TabIndex        =   27
               Top             =   360
               Width           =   1335
            End
            Begin VB.TextBox Text2 
               Height          =   285
               Index           =   7
               Left            =   6450
               MaxLength       =   2
               TabIndex        =   28
               Top             =   390
               Width           =   465
            End
            Begin VB.TextBox Text2 
               Height          =   300
               Index           =   8
               Left            =   7200
               MaxLength       =   10
               TabIndex        =   30
               Top             =   390
               Width           =   960
            End
            Begin VB.CommandButton cayuda 
               Caption         =   "..."
               Height          =   315
               Index           =   2
               Left            =   1260
               TabIndex        =   19
               Top             =   390
               Width           =   255
            End
            Begin VB.CommandButton cayuda 
               Caption         =   "..."
               Height          =   315
               Index           =   3
               Left            =   2700
               TabIndex        =   21
               Top             =   390
               Width           =   255
            End
            Begin VB.CommandButton cayuda 
               Caption         =   "..."
               Height          =   315
               Index           =   4
               Left            =   3930
               TabIndex        =   24
               Top             =   390
               Width           =   255
            End
            Begin VB.CommandButton cayuda 
               Caption         =   "..."
               Enabled         =   0   'False
               Height          =   315
               Index           =   5
               Left            =   4740
               TabIndex        =   26
               Top             =   390
               Width           =   255
            End
            Begin VB.CommandButton cayuda 
               Caption         =   "..."
               Height          =   285
               Index           =   6
               Left            =   6930
               TabIndex        =   29
               Top             =   390
               Width           =   195
            End
            Begin VB.TextBox Text2 
               Enabled         =   0   'False
               Height          =   285
               Index           =   9
               Left            =   120
               MaxLength       =   30
               TabIndex        =   32
               Top             =   1260
               Width           =   2715
            End
            Begin VB.CommandButton cayuda 
               Caption         =   "..."
               Enabled         =   0   'False
               Height          =   285
               Index           =   7
               Left            =   2850
               TabIndex        =   33
               Top             =   1260
               Width           =   195
            End
            Begin VB.TextBox Text2 
               Height          =   285
               Index           =   10
               Left            =   3210
               MaxLength       =   50
               TabIndex        =   34
               Top             =   1260
               Width           =   6105
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
               Left            =   8160
               TabIndex        =   31
               Top             =   390
               Width           =   1155
               _ExtentX        =   2037
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   10
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin VB.Label lblMonProv 
               BorderStyle     =   1  'Fixed Single
               Height          =   270
               Left            =   165
               TabIndex        =   68
               Top             =   720
               Width           =   885
            End
            Begin VB.Label Label6 
               BorderStyle     =   1  'Fixed Single
               Height          =   300
               Left            =   2700
               TabIndex        =   67
               Top             =   750
               Width           =   6615
            End
            Begin VB.Label Label3 
               Alignment       =   2  'Center
               Caption         =   "Item"
               Height          =   180
               Index           =   0
               Left            =   90
               TabIndex        =   65
               Top             =   150
               Width           =   645
            End
            Begin VB.Label Label3 
               Alignment       =   2  'Center
               Caption         =   "Tipo"
               Height          =   180
               Index           =   1
               Left            =   780
               TabIndex        =   64
               Top             =   150
               Width           =   405
            End
            Begin VB.Label Label3 
               Alignment       =   2  'Center
               Caption         =   "Numero"
               Height          =   180
               Index           =   2
               Left            =   1590
               TabIndex        =   63
               Top             =   150
               Width           =   1065
            End
            Begin VB.Label Label3 
               Alignment       =   2  'Center
               Caption         =   "T/P"
               Height          =   180
               Index           =   3
               Left            =   2940
               TabIndex        =   62
               Top             =   150
               Width           =   465
            End
            Begin VB.Label Label3 
               Alignment       =   2  'Center
               Caption         =   "T.Canc."
               Height          =   180
               Index           =   4
               Left            =   3420
               TabIndex        =   61
               Top             =   150
               Width           =   645
            End
            Begin VB.Label Label3 
               Alignment       =   2  'Center
               Caption         =   "Banco"
               Height          =   180
               Index           =   5
               Left            =   4260
               TabIndex        =   60
               Top             =   150
               Width           =   615
            End
            Begin VB.Label Label3 
               Alignment       =   2  'Center
               Caption         =   "Numero"
               Height          =   180
               Index           =   6
               Left            =   5130
               TabIndex        =   59
               Top             =   150
               Width           =   1245
            End
            Begin VB.Label Label3 
               Alignment       =   2  'Center
               Caption         =   "Mon."
               Height          =   180
               Index           =   7
               Left            =   6570
               TabIndex        =   58
               Top             =   150
               Width           =   465
            End
            Begin VB.Label Label3 
               Alignment       =   2  'Center
               Caption         =   "Importe"
               Height          =   210
               Index           =   8
               Left            =   7260
               TabIndex        =   57
               Top             =   150
               Width           =   765
            End
            Begin VB.Label Label3 
               Alignment       =   2  'Center
               Caption         =   "Fec. Cancela"
               Height          =   210
               Index           =   9
               Left            =   8280
               TabIndex        =   56
               Top             =   150
               Width           =   975
            End
            Begin VB.Label Label3 
               Alignment       =   2  'Center
               Caption         =   "Nro. Cuenta Corriente"
               Height          =   180
               Index           =   10
               Left            =   180
               TabIndex        =   55
               Top             =   1020
               Width           =   2625
            End
            Begin VB.Label Label3 
               Alignment       =   2  'Center
               Caption         =   "Observaciones"
               Height          =   180
               Index           =   11
               Left            =   3240
               TabIndex        =   54
               Top             =   1050
               Width           =   6075
            End
         End
      End
      Begin VB.Frame Frame2 
         Height          =   2520
         Left            =   180
         TabIndex        =   13
         Top             =   240
         Width           =   9645
         Begin VB.ComboBox Combo1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1620
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   180
            Width           =   2235
         End
         Begin VB.TextBox Text1 
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
            Left            =   5490
            MaxLength       =   6
            TabIndex        =   1
            Top             =   210
            Width           =   1125
         End
         Begin VB.ComboBox Combo2 
            Height          =   315
            Left            =   1620
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   2085
            Width           =   1965
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   3
            Left            =   4740
            MaxLength       =   10
            TabIndex        =   9
            Top             =   2130
            Width           =   855
         End
         Begin VB.TextBox Text1 
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
            Left            =   1620
            MaxLength       =   2
            TabIndex        =   5
            Top             =   1350
            Width           =   465
         End
         Begin VB.TextBox Text1 
            Enabled         =   0   'False
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
            Index           =   2
            Left            =   9045
            MaxLength       =   2
            TabIndex        =   10
            Top             =   1335
            Width           =   465
         End
         Begin VB.CommandButton cayuda 
            Caption         =   "..."
            Height          =   315
            Index           =   0
            Left            =   2100
            TabIndex        =   6
            Top             =   1305
            Width           =   225
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H0080C0FF&
            BorderStyle     =   0  'None
            Height          =   765
            Left            =   5640
            TabIndex        =   14
            Top             =   1665
            Width           =   3915
            Begin VB.Label Label4 
               Alignment       =   2  'Center
               BackColor       =   &H00004000&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "TOTALES"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H0080FFFF&
               Height          =   390
               Index           =   0
               Left            =   0
               TabIndex        =   37
               Top             =   0
               Width           =   3915
            End
            Begin VB.Label Label4 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "S/."
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   1
               Left            =   120
               TabIndex        =   35
               Top             =   480
               Width           =   345
            End
            Begin VB.Label Label4 
               Alignment       =   2  'Center
               BackColor       =   &H00C0FFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "US$"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   2
               Left            =   1950
               TabIndex        =   17
               Top             =   480
               Width           =   435
            End
            Begin VB.Label Label5 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               BorderStyle     =   1  'Fixed Single
               Enabled         =   0   'False
               Height          =   285
               Index           =   0
               Left            =   480
               TabIndex        =   16
               Top             =   420
               Width           =   1365
            End
            Begin VB.Label Label5 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               BorderStyle     =   1  'Fixed Single
               Enabled         =   0   'False
               Height          =   285
               Index           =   1
               Left            =   2430
               TabIndex        =   15
               Top             =   420
               Width           =   1365
            End
         End
         Begin MSMask.MaskEdBox MBox1 
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
            Left            =   8250
            TabIndex        =   2
            Top             =   210
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuda2 
            Height          =   345
            Left            =   1620
            TabIndex        =   3
            Top             =   600
            Width           =   7875
            _ExtentX        =   13891
            _ExtentY        =   609
            XcodMaxLongitud =   11
            xcodwith        =   800
            NomTabla        =   "vt_cliente"
            TituloAyuda     =   "Ayuda de Clientes"
            ListaCampos     =   "clientecodigo(1),clienterazonsocial(1)"
            XcodCampo       =   "clientecodigo"
            XListCampo      =   "clienterazonsocial"
            ListaCamposDescrip=   "Codigo,Descripcion"
            ListaCamposText =   "clientecodigo,clienterazonsocial"
         End
         Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyudaCaja 
            Height          =   315
            Left            =   1620
            TabIndex        =   7
            Top             =   1680
            Width           =   3900
            _ExtentX        =   6879
            _ExtentY        =   556
            XcodMaxLongitud =   11
            xcodwith        =   400
            NomTabla        =   "te_codigocaja"
            TituloAyuda     =   "Busqueda de Caja"
            ListaCampos     =   "cajacodigo(1),cajadescripcion(1),cajarendiciones(2)"
            XcodCampo       =   "cajacodigo"
            XListCampo      =   "cajadescripcion"
            ListaCamposDescrip=   "Codigo,Descripcion,controla Rendicion"
            ListaCamposText =   "cajacodigo,cajadescripcion,cajarendiciones"
         End
         Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuempresa 
            Height          =   315
            Left            =   1620
            TabIndex        =   4
            Top             =   945
            Width           =   4050
            _ExtentX        =   7144
            _ExtentY        =   556
            XcodMaxLongitud =   3
            xcodwith        =   300
            NomTabla        =   "co_multiempresas"
            TituloAyuda     =   "Busqueda de Empresas"
            ListaCampos     =   "empresacodigo(1),empresadescripcion(1)"
            XcodCampo       =   "empresacodigo"
            XListCampo      =   "empresadescripcion"
            ListaCamposDescrip=   "Codigo,Descripcion"
            ListaCamposText =   "empresacodigo,empresadescripcion"
         End
         Begin VB.Label Lblempresa 
            AutoSize        =   -1  'True
            Caption         =   "Empresa :"
            Height          =   195
            Left            =   315
            TabIndex        =   69
            Top             =   1005
            Width           =   705
         End
         Begin VB.Label Label1 
            Caption         =   "Ingreso/Egreso"
            Height          =   255
            Index           =   0
            Left            =   300
            TabIndex        =   49
            Top             =   210
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "Recibo"
            Height          =   255
            Index           =   1
            Left            =   4470
            TabIndex        =   48
            Top             =   210
            Width           =   825
         End
         Begin VB.Label Label1 
            Caption         =   "Ingreso/Egreso"
            Height          =   255
            Index           =   2
            Left            =   6900
            TabIndex        =   47
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "Operacion"
            Height          =   255
            Index           =   3
            Left            =   300
            TabIndex        =   46
            Top             =   1335
            Width           =   1275
         End
         Begin VB.Label Label1 
            Caption         =   "Cliente"
            Height          =   255
            Index           =   4
            Left            =   300
            TabIndex        =   45
            Top             =   630
            Width           =   1035
         End
         Begin VB.Label Label1 
            Caption         =   "Cod. Caja"
            Height          =   255
            Index           =   5
            Left            =   300
            TabIndex        =   44
            Top             =   1695
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "Tipo Cambio"
            Height          =   255
            Index           =   6
            Left            =   3780
            TabIndex        =   43
            Top             =   1830
            Width           =   1005
         End
         Begin VB.Label Label1 
            Caption         =   "Moneda"
            Height          =   255
            Index           =   7
            Left            =   300
            TabIndex        =   42
            Top             =   2085
            Width           =   1455
         End
         Begin VB.Label Label2 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Index           =   0
            Left            =   2340
            TabIndex        =   40
            Top             =   1335
            Width           =   3255
         End
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00C0C0C0&
      Height          =   990
      Left            =   3480
      TabIndex        =   11
      Top             =   8055
      Width           =   4230
      Begin VB.CommandButton cmdBotones 
         Caption         =   "&Nuevo"
         Height          =   690
         Index           =   4
         Left            =   180
         Picture         =   "FrmMovimientoCliente.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   210
         Width           =   870
      End
      Begin VB.CommandButton cmdBotones 
         Caption         =   "&Grabar"
         Height          =   690
         Index           =   5
         Left            =   1200
         Picture         =   "FrmMovimientoCliente.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   180
         Width           =   870
      End
      Begin VB.CommandButton cmdBotones 
         Caption         =   "&Eliminar"
         Height          =   690
         Index           =   6
         Left            =   2250
         Picture         =   "FrmMovimientoCliente.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   210
         Width           =   825
      End
      Begin VB.CommandButton cmdBotones 
         Caption         =   "&Salir"
         Height          =   690
         Index           =   7
         Left            =   3255
         Picture         =   "FrmMovimientoCliente.frx":0CC6
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   210
         Width           =   870
      End
   End
   Begin MSComctlLib.StatusBar Panel 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   66
      Top             =   9105
      Width           =   10455
      _ExtentX        =   18441
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
Attribute VB_Name = "FrmMovimientoClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adll As New dllgeneral.dll_general
Dim rsdetat As New ADODB.Recordset

Dim m_fondofijo As Integer
Dim m_docxrendir As Integer
Property Let docxrendir(valor As String)
   m_docxrendir = valor
End Property
Property Let fondofijo(valor As String)
   m_fondofijo = valor
End Property
Public Sub ConfigGrid()
    With TDBGrid1
        .Columns(0).Width = 600
        .Columns(1).Width = 600
        .Columns(2).Width = 1500
        .Columns(3).Width = 600
        .Columns(4).Width = 700
        .Columns(5).Width = 700
        .Columns(6).Width = 1500
        .Columns(7).Width = 600
        .Columns(8).HeadAlignment = dbgCenter
        .Columns(8).Width = 1300
        .Columns(8).NumberFormat = "##,###,##0.00"
        .Columns(9).Width = 1000
        .Columns(10).Width = 1500
        .Columns(11).Width = 2000
        .Refresh
    End With
End Sub
Public Sub cargar_grilla()
   Set rsdetat = Nothing
   Call rsdetat.Fields.Append("Item", adChar, 3)
   Call rsdetat.Fields.Append("Tipo", adChar, 2)
   Call rsdetat.Fields.Append("Numero", adChar, 14)
   Call rsdetat.Fields.Append("T/P", adChar, 1)
   Call rsdetat.Fields.Append("T.Canc", adChar, 2)
   Call rsdetat.Fields.Append("Banco", adChar, 2)
   Call rsdetat.Fields.Append("Numero Doc", adChar, 20)
   Call rsdetat.Fields.Append("Mnda", adChar, 2)
   Call rsdetat.Fields.Append("Importe", adDouble)
   Call rsdetat.Fields.Append("Fecha Canc", adDate)
   Call rsdetat.Fields.Append("Cta Cte", adChar, 30)
   Call rsdetat.Fields.Append("Observaciones", adChar, 50)
   
   rsdetat.Open
   Set TDBGrid1.DataSource = rsdetat
   TDBGrid1.Refresh
   Call ConfigGrid
   
End Sub

Private Sub cAyuda_Click(Index As Integer)
 Dim rb As New ADODB.Recordset
 Dim nMonedaCab As String
 nAyuda = "": nDetalle = ""
 nAyuda1 = ""
 nMoneda = ""
  If Index = 0 Then
         If Len(Trim(Text1(1))) > 0 Then
           SendKeys "{tab}"
           Exit Sub
         End If
         Dim dfiltra(1, 2) As String
         dfiltra(1, 1) = "Codigo": dfiltra(1, 2) = "operacioncodigo"
         FrmAyudaTes.TipoForma = 1
         FrmAyudaTes.BConexion = VGCNx
         FrmAyudaTes.Bdata = "0"
         FrmAyudaTes.BTabla = "te_operaciongeneral"
         FrmAyudaTes.BCampos = "operacioncodigo as Codigo,operaciondescripcion as Descripcion"
         FrmAyudaTes.BOrden = "operacioncodigo"
         FrmAyudaTes.BCondi = "operacioncontrolaclienteprov='" & IIf(adll.ComboDato(Combo1) = "I", "C", "C") & "'"
         FrmAyudaTes.BFiltro = dfiltra
         FrmAyudaTes.Show 1
         Text1(1) = nAyuda
         Label2(0) = nDetalle
         Call Text1_KeyPress(1, 13)
         
    ElseIf Index = 1 Then
         Set rb = VGCNx.Execute("select * from te_operaciongeneral where operacioncodigo='" & Escadena(Text1(1)) & "' and operacioncontrolaclienteprov='" & IIf(adll.ComboDato(Combo1) = "I", "C", "C") & "'")
         If rb.RecordCount > 0 Then
            If Escadena(rb!operacionvalidacajabancos) = "B" Then
                Text1(2).Enabled = False
                cayuda(1).Enabled = False
                Combo2.SetFocus
                rb.Close
                Set rb = Nothing
                Exit Sub
            Else
'                Text1(2).Enabled = True
                cayuda(1).Enabled = True
'                Text1(2).SetFocus
            End If
        End If
        rb.Close
        Set rb = Nothing
         
        If Len(Trim(Text1(2))) > 0 Then
           SendKeys "{tab}"
           Exit Sub
        End If
        
        Dim gfiltra(1, 2) As String
        gfiltra(1, 1) = "Codigo": gfiltra(1, 2) = "cajacodigo"
        FrmAyudaTes.TipoForma = 1
        FrmAyudaTes.BConexion = VGCNx
        FrmAyudaTes.Bdata = "0"
        FrmAyudaTes.BTabla = "te_codigocaja"
        FrmAyudaTes.BCampos = "cajacodigo as Codigo,cajadescripcion as Descripcion"
        FrmAyudaTes.BOrden = "cajacodigo"
        FrmAyudaTes.BCondi = ""
        FrmAyudaTes.BFiltro = gfiltra
        FrmAyudaTes.Show 1
        Text1(2) = nAyuda
        Label2(1) = nDetalle
        SendKeys "{tab}"
     ElseIf Index = 2 Then
       If Len(Trim(Text2(1))) > 0 Then
          SendKeys "{tab}"
          Exit Sub
        End If
        If adll.VerificaDatoExistente(VGCNx, "select * from cc_tipodocumento where tdocumentoingplan='1'") = 1 Then
            Dim zfiltra(1, 2) As String
            zfiltra(1, 1) = "Documento": zfiltra(1, 2) = "tdocumentocodigo"
            FrmAyudaTes.TipoForma = 1
            FrmAyudaTes.BConexion = VGCNx
            FrmAyudaTes.Bdata = "0"
            FrmAyudaTes.BTabla = "cc_tipodocumento"
            FrmAyudaTes.BCampos = "tdocumentocodigo as Codigo,tdocumentodescripcion as Descripcion"
            FrmAyudaTes.BOrden = "tdocumentocodigo"
            FrmAyudaTes.BCondi = "tdocumentoingplan='1'"   'tdocumentotipo='C'"
            FrmAyudaTes.BFiltro = zfiltra
            FrmAyudaTes.Show 1
            Text2(1) = nAyuda
            Call Text2_KeyPress(1, 13)
         Else
             nAyuda = "": nDetalle = ""
             MsgBox "No existen Documentos...", vbInformation, MsgTitle
             Exit Sub
         End If
    ElseIf Index = 3 Then
         If Len(Trim(Text2(2))) > 0 Then
           SendKeys "{tab}"
           Exit Sub
         End If
         If adll.VerificaDatoExistente(VGCNx, "select * from vt_cargo where empresacodigo ='" & Ctr_Ayuempresa.xclave & "' and clientecodigo='" & Ctr_Ayuda2.xclave & "' and documentocargo='" & Text2(1) & "'") = 1 Then
            Dim wfiltra(1, 2) As String
            wfiltra(1, 1) = "Documento": wfiltra(1, 2) = "cargonumdoc"
            FrmAyudaTes.TipoForma = 5
            FrmAyudaTes.BConexion = VGCNx
            FrmAyudaTes.Bdata = "0"
            FrmAyudaTes.BTabla = "vt_cargo A inner join cc_tipodocumento B On a.documentocargo=b.tdocumentocodigo"
            FrmAyudaTes.BCampos = "documentocargo as TD,cargonumdoc as Numero,monedacodigo as Mnd,cargoapeimpape as Total,(Round(cargoapeimpape,2)-Round(isnull(cargoapeimppag,0),2)) as Saldo,cargoapefecemi as FecEmision,cargoapefecvct as FecVencimiento, cargoaperefere as Referencia"
            FrmAyudaTes.BOrden = "Clientecodigo,cargoapefecemi"
            FrmAyudaTes.BCondi = " empresacodigo ='" & Ctr_Ayuempresa.xclave & "' and clientecodigo='" & Ctr_Ayuda2.xclave & "' and isnull(cargoapeflgcan,0)<>1  and b.tdocumentotipo='C' and a.documentocargo='" & Text2(1).Text & "' and isnull(cargoapeflgreg,0)<>1"
            FrmAyudaTes.BFiltro = gfiltra
            FrmAyudaTes.Show 1
            Text2(2) = nDetalle
            Text2(8).Text = nAyuda
            Label6.Caption = nAyuda1
            If nAyuda = Empty Then Exit Sub
            lblMonProv.Caption = nMoneda
            nMonedaCab = Left(Combo2.List(Combo2.ListIndex), 2)
            Text2(7).Text = nMonedaCab
            If lblMonProv.Caption <> nMonedaCab Then
               If nMonedaCab = "01" Then
                  Text2(8).Text = Format(nAyuda * Text1(3).Text, "###,###,##0.#0")
               Else
                  Text2(8).Text = Format(nAyuda / Text1(3).Text, "###,###,##0.#0")
               End If
            End If
            
         Else
            nAyuda = "": nDetalle = ""
            MsgBox "No existen Documentos...", vbInformation, MsgTitle
            Exit Sub
         End If
  ElseIf Index = 4 Then   'Tipo de cancelacion
    If Len(Trim(Text2(4))) > 0 Then
      SendKeys "{tab}"
      Exit Sub
    End If
    If adll.VerificaDatoExistente(VGCNx, "select * from cc_tipodocumento where tdocumentotipo='A' and tdocumentoingcobra='1'") = 1 Then
        Dim ffiltra(1, 2) As String
        ffiltra(1, 1) = "Documento": ffiltra(1, 2) = "tdocumentocodigo"
        FrmAyudaTes.TipoForma = 1
        FrmAyudaTes.BConexion = VGCNx
        FrmAyudaTes.Bdata = "0"
        FrmAyudaTes.BTabla = "cc_tipodocumento"
        FrmAyudaTes.BCampos = "tdocumentocodigo as Codigo,tdocumentodescripcion as Descripcion"
        FrmAyudaTes.BOrden = "tdocumentocodigo"
        FrmAyudaTes.BCondi = "tdocumentotipo='A' and tdocumentocancela='1'"
        FrmAyudaTes.BFiltro = ffiltra
        FrmAyudaTes.Show 1
        Text2(4).Text = nAyuda
        
        If Trim(nAyuda) = "10" Then
           Text2(6).Enabled = False
           cayuda(5).Enabled = False
           Text2(5).Enabled = False
        Else
           Text2(6).Enabled = True
           cayuda(5).Enabled = True
           Text2(5).Enabled = True
        End If
        
     Else
         nAyuda = "": nDetalle = ""
         MsgBox "No existen Documentos...", vbInformation, MsgTitle
         Exit Sub
     End If
     Exit Sub
   ElseIf Index = 5 Then    'Tipo de Banco
        If Len(Trim(Text2(5))) > 0 Then
          SendKeys "{tab}"
          Exit Sub
        End If
        Dim tfiltra(1, 2) As String
        tfiltra(1, 1) = "Banco": tfiltra(1, 2) = "bancodescripcion"
        FrmAyudaTes.TipoForma = 1
        FrmAyudaTes.BConexion = VGCNx
        FrmAyudaTes.Bdata = "0"
        FrmAyudaTes.BTabla = "gr_banco a INNER JOIN te_cuentabancos b ON a.bancocodigo=b.cbanco_codigo"
        FrmAyudaTes.BCampos = "DISTINCT a.bancocodigo as Codigo,a.bancodescripcion as Descripcion"
        FrmAyudaTes.BOrden = "a.bancocodigo"
        FrmAyudaTes.BCondi = "b.empresacodigo='" & Ctr_Ayuempresa.xclave & "'"
        FrmAyudaTes.BFiltro = tfiltra
        FrmAyudaTes.Show 1
        Text2(5) = nAyuda
   ElseIf Index = 6 Then    'Tipo de Moneda
        If Len(Trim(Text2(7))) > 0 Then
           SendKeys "{tab}"
           Exit Sub
        End If
        Dim pfiltra(1, 2) As String
        pfiltra(1, 1) = "Codigo": pfiltra(1, 2) = "monedacodigo"
        FrmAyudaTes.TipoForma = 1
        FrmAyudaTes.BConexion = VGCNx
        FrmAyudaTes.Bdata = "0"
        FrmAyudaTes.BTabla = "gr_moneda"
        FrmAyudaTes.BCampos = "monedacodigo as Codigo,monedadescripcion as Descripcion"
        FrmAyudaTes.BOrden = "monedacodigo"
        FrmAyudaTes.BCondi = ""
        FrmAyudaTes.BFiltro = pfiltra
        FrmAyudaTes.Show 1
        Text2(7) = nAyuda
   ElseIf Index = 7 Then    'Nro Cuenta Corriente
        If Len(Trim(Text2(9))) > 0 Then
           SendKeys "{tab}"
           Exit Sub
        End If
        Dim qfiltra(1, 2) As String
        qfiltra(1, 1) = "Banco": qfiltra(1, 2) = "bancocodigo"
        FrmAyudaTes.TipoForma = 1
        FrmAyudaTes.BConexion = VGCNx
        FrmAyudaTes.Bdata = "0"
        FrmAyudaTes.BTabla = "te_cuentabancos inner join gr_banco on te_cuentabancos.cbanco_codigo=gr_banco.bancocodigo"
        FrmAyudaTes.BCampos = "cbanco_numero as NoCtaCte,monedacodigo as Moneda,bancocodigo as CodBan,bancodescripcion as Banco"
        FrmAyudaTes.BOrden = "gr_banco.bancocodigo"
        FrmAyudaTes.BCondi = "gr_banco.bancocodigo='" & Text2(5) & "' and te_cuentabancos.monedacodigo='" & adll.ComboDato(Combo2) & "' and te_cuentabancos.empresacodigo='" & Ctr_Ayuempresa.xclave & "'"
        FrmAyudaTes.BFiltro = qfiltra
        FrmAyudaTes.Show 1
        Text2(9) = nAyuda
        
   End If
   nAyuda = "": nDetalle = ""
End Sub

Public Function GrabarData() As Integer
  Dim acmd As New ADODB.Command
  Dim rb As New ADODB.Recordset
  Dim ingresacargo As Integer
  Dim xabono, xzona, xmone, xcuenta, xtipo As String
  Dim xnumplan, ximpsol, xtcam, xnumpag As Double
 On Error GoTo error
    GrabarData = 0
    ingresacargo = 0
VGCNx.BeginTrans
    'Actualizamos el numerador de tipo de ingreso
    Set rb = VGCNx.Execute("select * from te_parametroempresa where empresacodigo='" & VGCodEmpresa & "'")
    If rb.RecordCount > 0 Then
     If adll.ComboDato(Combo1.Text) = "I" Then
         Text1(0) = Right("0000000000" & Trim(CStr(CDbl(IIf(IsNull(rb!empresanumeingreso + 1) Or Len(Trim(rb!empresanumeingreso)) = 0, 1, rb!empresanumeingreso + 1)))), 6)
         VGCNx.Execute "Update te_parametroempresa Set empresanumeingreso='" & Right("0000000000" & Trim(CStr(Val(Text1(0)))), 6) & "' where empresacodigo='" & VGCodEmpresa & "'"
         
     ElseIf adll.ComboDato(Combo1.Text) = "E" Then
         Text1(0) = Right("0000000000" & Trim(CStr(CDbl(IIf(IsNull(rb!empresanumegreso + 1) Or Len(Trim(rb!empresanumegreso)) = 0, 1, rb!empresanumegreso + 1)))), 6)
         VGCNx.Execute "Update te_parametroempresa Set empresanumegreso='" & Right("0000000000" & Trim(CStr(Val(Text1(0)))), 6) & "' where empresacodigo='" & VGCodEmpresa & "'"
     End If
    End If
    rb.Close
    Set rb = Nothing
VGCNx.CommitTrans
VGCNx.BeginTrans
    Set acmd.ActiveConnection = VGGeneral
    acmd.CommandType = adCmdStoredProc
    acmd.CommandText = "te_abonadocumento_pro"
    acmd.CommandTimeout = 0
    acmd.Prepared = True
    With acmd
        .Parameters("@base") = VGCNx.DefaultDatabase
        .Parameters("@tipo") = "1"
        .Parameters("@numrecibo") = Escadena(Text1(0))
        .Parameters("@estadoreg") = ""
        .Parameters("@controlctacte") = "1"
        .Parameters("@vendedorcodigo") = VGoficina
        .Parameters("@cajacodigo") = Trim(Ctr_AyudaCaja.xclave)
        .Parameters("@clientecodigo") = Escadena(Ctr_Ayuda2.xclave)
        .Parameters("@descripcion") = ""
        .Parameters("@operacion") = Escadena(Text1(1))
        .Parameters("@monedacodigo") = adll.ComboDato(Combo2)
        .Parameters("@ingsal") = adll.ComboDato(Combo1)
        .Parameters("@tipocambio") = CDbl(Text1(3))
        .Parameters("@totsoles") = CDbl(Label5(0))
        .Parameters("@totdolares") = CDbl(Label5(1))
        .Parameters("@fechadocumento") = MBox1.Text
        .Parameters("@empresa") = Ctr_Ayuempresa.xclave
        .Parameters("@observa") = ""
        .Parameters("@transferauto") = ""
        .Parameters("@numreciboegreso") = ""
        .Parameters("@usuario") = VGUsuario
        .Parameters("@fechaact") = Now
     End With
     acmd.Execute
     Set acmd = Nothing
     xmone = adll.ComboDato(Combo2)
     If rsdetat.RecordCount > 0 Then
         rsdetat.MoveLast
         rsdetat.MoveFirst
         Do Until rsdetat.EOF
             Set rb = VGCNx.Execute("select * from cc_tipodocumento where tdocumentocodigo='" & rsdetat.Fields(1) & "'")
     '         ingresacargo = ESNULO(rb!tdocumentotipo, "")
             If adll.ComboDato(Combo1) = "E" And ESNULO(rb!tdocumentotipo, "") = "C" Then ingresacargo = 1
              Set rb = VGCNx.Execute("select * from cc_tipodocumento where tdocumentocodigo='" & rsdetat.Fields(4) & "'")
             xzona = "01": xnumpag = 1
             If rb.RecordCount > 0 Then
                xabono = rb!tdocumentotipo
                xtipo = IIf(IsNull(rb!tdocumentopermiteaplica), Null, rb!tdocumentopermiteaplica)
                If rsdetat.Fields(7) = g_tiposol Then
                   xcuenta = "" & Trim(rb!tdocumentocuentasoles)
                Else
                   xcuenta = "" & Trim(rb!tdocumentocuentadolares)
                End If
             Else
                xabono = "": xcuenta = "": xtipo = ""
             End If
             rb.Close
             Set rb = Nothing
        
             Set rb = VGCNx.Execute("select * from vt_cargo where documentocargo='" & rsdetat.Fields(1) & "' and cargonumdoc='" & rsdetat.Fields(2) & "'")
             If rb.RecordCount > 0 Then
                xzona = rb!zonacodigo
                xmone = rb!monedacodigo
                If IsNull(rb!cargoapenumpag) Then
                  xnumpag = 1
                Else
                  xnumpag = Val(rb!cargoapenumpag)
                End If
             Else
                xzona = "01": xnumpag = 1
             End If
             rb.Close
             Set rb = Nothing
             
             ximpsol = CDbl(rsdetat.Fields(8))
             xtcam = CDbl(Text1(3))
             If rsdetat.Fields(7) <> xmone Then
                If rsdetat.Fields(7) = g_tiposol Then
                   xtcam = CDbl(Text1(3))
                   If CDbl(Text1(3)) = 0 Or Len(Trim(Text1(3))) = 0 Then xtcam = 1
                   ximpsol = CDbl(rsdetat.Fields(8)) / CDbl(xtcam)
                Else
                   xtcam = CDbl(Text1(3))
                   If CDbl(Text1(3)) = 0 Or Len(Trim(Text1(3))) = 0 Then xtcam = 1
                   ximpsol = CDbl(rsdetat.Fields(8)) * CDbl(xtcam)
                End If
             End If
          If ingresacargo = 0 Then
             Set acmd.ActiveConnection = VGGeneral
             acmd.CommandType = adCmdStoredProc
             acmd.CommandText = "cc_abonadocumento_pro"
             acmd.CommandTimeout = 0
             acmd.Prepared = True
             With acmd
                 .Parameters("@base") = VGCNx.DefaultDatabase
                 .Parameters("@tipo") = "1"
                 .Parameters("@documentoabono") = rsdetat.Fields(1)
                 .Parameters("@abononumdoc") = Trim(rsdetat.Fields(2))
                 .Parameters("@abonocannumpag") = xnumpag
                 .Parameters("@zonacodigo") = xzona
                 .Parameters("@tipoplanilla") = "TE"
                 .Parameters("@vendedor") = ""    'Escadena(Ctr_Ayuda2.xclave)
                 .Parameters("@numplanilla") = Right("00000000" & Trim(Text1(0)), 6)
                 .Parameters("@fechapla") = MBox1.Text
                 .Parameters("@fechapro") = MBox1.Text
                 .Parameters("@moneda") = xmone
                 .Parameters("@abonocancarabo") = xabono
                 .Parameters("@cuenta") = xcuenta
                 .Parameters("@banco") = "" & Trim(rsdetat.Fields(5))
                 .Parameters("@tipocam") = CDbl(xtcam)
                 .Parameters("@abonoflpres") = "1"
                 .Parameters("@abonocanimpcan") = CDbl(rsdetat.Fields(8))
                 .Parameters("@abonocanimpsol") = ximpsol
                 .Parameters("@usuario") = VGUsuario
                 .Parameters("@fechaact") = Now
                 .Parameters("@forma") = rsdetat.Fields(3)
                 .Parameters("@monedacan") = adll.ComboDato(Combo2)
                 .Parameters("@abonocantd") = rsdetat.Fields(4)
                 .Parameters("@abonocannro") = Trim(rsdetat.Fields(6))
                 .Parameters("@fechacan") = rsdetat.Fields(9)
                 .Parameters("@cliente") = Escadena(Ctr_Ayuda2.xclave)
                 .Parameters("@empresa") = Ctr_Ayuempresa.xclave
             End With
             acmd.Execute
             
             Set acmd = Nothing
             DoEvents
   
             '**** Actualizamos Saldos de documento pendiente
             If rsdetat.Fields(7) = g_tipodolar Then
                If xmone = g_tiposol Then
                    VGCNx.Execute "Update  vt_cargo Set cargoapeimppag=isnull(cargoapeimppag,0)+" & CDbl(rsdetat.Fields(8) * xtcam) & "," & _
                               " cargoapenumpag='" & xnumpag + 1 & "'" & _
                               " Where documentocargo='" & rsdetat.Fields(1) & "' and cargonumdoc='" & Trim(rsdetat.Fields(2)) & "' and " & _
                               " clientecodigo='" & Trim(Ctr_Ayuda2.xclave) & "'"
                Else
                    VGCNx.Execute "Update  vt_cargo Set cargoapeimppag=isnull(cargoapeimppag,0)+" & CDbl(rsdetat.Fields(8)) & "," & _
                               " cargoapenumpag='" & xnumpag + 1 & "'" & _
                               " Where documentocargo='" & rsdetat.Fields(1) & "' and cargonumdoc='" & Trim(rsdetat.Fields(2)) & "' and " & _
                               " clientecodigo='" & Trim(Ctr_Ayuda2.xclave) & "'"
                End If
             ElseIf rsdetat.Fields(7) = g_tiposol Then
                If xmone = g_tipodolar Then
                    VGCNx.Execute "Update  vt_cargo Set cargoapeimppag=isnull(cargoapeimppag,0)+" & CDbl(rsdetat.Fields(8) / xtcam) & "," & _
                               " cargoapenumpag='" & xnumpag + 1 & "'" & _
                               " Where documentocargo='" & rsdetat.Fields(1) & "' and cargonumdoc='" & Trim(rsdetat.Fields(2)) & "' and " & _
                               " clientecodigo='" & Trim(Ctr_Ayuda2.xclave) & "'"
                Else
                    VGCNx.Execute "Update  vt_cargo Set cargoapeimppag=isnull(cargoapeimppag,0)+" & CDbl(rsdetat.Fields(8)) & "," & _
                               " cargoapenumpag='" & xnumpag + 1 & "'" & _
                               " Where documentocargo='" & rsdetat.Fields(1) & "' and cargonumdoc='" & Trim(rsdetat.Fields(2)) & "' and " & _
                               " clientecodigo='" & Trim(Ctr_Ayuda2.xclave) & "'"
                End If
             End If
             
             VGCNx.Execute "Update vt_cargo " & _
                         " Set cargoapeflgcan= CASE Round(isnull(cargoapeimpape,0),2)-Round(isnull(cargoapeimppag,0),2) WHEN 0 THEN '1' ELSE '0' END ," & _
                         "   cargoapefeccan='" & rsdetat.Fields(9) & "'" & _
                         " Where documentocargo='" & rsdetat.Fields(1) & "' and cargonumdoc='" & Trim(rsdetat.Fields(2)) & "' and " & _
                         " clientecodigo='" & Trim(Ctr_Ayuda2.xclave) & "'"
             
             '**** Actualizamos Saldos del cliente
             If rsdetat.Fields(7) = g_tipodolar Then
                   VGCNx.Execute "Update  vt_cliente Set clientesaldodolares=isnull(clientesaldodolares,0)-" & CDbl(rsdetat.Fields(8)) & _
                               " Where clientecodigo='" & Escadena(Ctr_Ayuda2.xclave) & "'"
             ElseIf rsdetat.Fields(7) = g_tiposol Then
                   VGCNx.Execute "Update  vt_cliente Set clientesaldosoles=isnull(clientesaldosoles,0)-" & CDbl(rsdetat.Fields(8)) & _
                               " Where clientecodigo='" & Escadena(Ctr_Ayuda2.xclave) & "'"
             End If
                                             
            '**** Actualizamos correlativo de documentos de anticipos
            Dim rsql As String
             rsql = " select tdocumentonumeauto,tdocumentonumerador from cc_tipodocumento "
             rsql = rsql & " where tdocumentocodigo='" & Trim(rsdetat.Fields(1)) & "'"
             Set rb = VGCNx.Execute(rsql)
             If rb!tdocumentonumeauto = 1 And rb!tdocumentonumerador = rsdetat.Fields(2) Then
                rsdetat.Fields(2) = rb!tdocumentonumerador
                rsql = Format(rsdetat.Fields(2) + 1, "00000000000000")
                VGCNx.Execute "Update  cc_tipodocumento Set tdocumentonumerador='" & rsql & "' where tdocumentocodigo='" & Trim(rsdetat.Fields(1)) & "'"
             End If
             
             '****Permite Aplicaciones
             If Not IsNull(xtipo) Then
                 If xtipo = 1 Then
                         Set rb = VGCNx.Execute("select * from cc_tipodocumento where tdocumentocodigo='" & rsdetat.Fields(1) & "'")
                         If rb.RecordCount > 0 Then
                            xabono = rb!tdocumentotipo
                            If rsdetat.Fields(7) = g_tiposol Then
                               xcuenta = "" & Trim(rb!tdocumentocuentasoles)
                            Else
                               xcuenta = "" & Trim(rb!tdocumentocuentadolares)
                            End If
                         Else
                            xabono = "": xcuenta = ""
                         End If
                         rb.Close
                         Set rb = Nothing
                         
                         Set rb = VGCNx.Execute("select * from vt_cargo where documentocargo='" & rsdetat.Fields(4) & "' and cargonumdoc='" & rsdetat.Fields(6) & "'")
                         If rb.RecordCount > 0 Then
                            xzona = rb!zonacodigo
                            xmone = rb!monedacodigo
                            If IsNull(rb!cargoapenumpag) Then
                              xnumpag = 1
                            Else
                              xnumpag = Val(rb!cargoapenumpag)
                            End If
                         Else
                            xzona = "01":  xnumpag = 1
                         End If
                         rb.Close
                         Set rb = Nothing
                                                                     
                         ximpsol = CDbl(rsdetat.Fields(8))
                         xtcam = CDbl(Text1(3))
                         If rsdetat.Fields(7) <> xmone Then
                            If rsdetat.Fields(7) = g_tiposol Then
                               xtcam = CDbl(Text1(3))
                               If CDbl(Text1(3)) = 0 Or Len(Trim(Text1(3))) = 0 Then xtcam = 1
                               ximpsol = CDbl(rsdetat.Fields(8)) / CDbl(xtcam)
                            Else
                               xtcam = CDbl(Text1(3))
                               If CDbl(Text1(3)) = 0 Or Len(Trim(Text1(3))) = 0 Then xtcam = 1
                                ximpsol = CDbl(rsdetat.Fields(8)) * CDbl(xtcam)
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
                             .Parameters("@documentoabono") = rsdetat.Fields(4)
                             .Parameters("@abononumdoc") = Trim(rsdetat.Fields(6))
                             .Parameters("@abonocannumpag") = xnumpag
                             .Parameters("@zonacodigo") = xzona
                             .Parameters("@tipoplanilla") = "TE" ' Escadena(Ctr_Ayuda1.xclave)
                             .Parameters("@vendedor") = ""  'Escadena(Ctr_Ayuda2.xclave)
                             .Parameters("@numplanilla") = Right("00000000" & Trim(Text1(0)), 6)
                             .Parameters("@fechapla") = MBox1.Text
                             .Parameters("@fechapro") = MBox1.Text
                             .Parameters("@moneda") = xmone
                             .Parameters("@abonocancarabo") = "A"   'xabono
                             .Parameters("@cuenta") = xcuenta
                             .Parameters("@banco") = "" & Trim(rsdetat.Fields(5))
                             .Parameters("@tipocam") = CDbl(xtcam)
                             .Parameters("@abonoflpres") = "1"
                             .Parameters("@abonocanimpcan") = CDbl(rsdetat.Fields(8))
                             .Parameters("@abonocanimpsol") = ximpsol
                             .Parameters("@usuario") = VGUsuario
                             .Parameters("@fechaact") = Now
                             .Parameters("@forma") = rsdetat.Fields(3)
                             .Parameters("@monedacan") = rsdetat.Fields(7)
                             .Parameters("@abonocantd") = rsdetat.Fields(1)
                             .Parameters("@abonocannro") = Trim(rsdetat.Fields(2))
                             .Parameters("@fechacan") = rsdetat.Fields(9)
                             .Parameters("@cliente") = Escadena(Ctr_Ayuda2.xclave)
                             .Parameters("@empresa") = Ctr_Ayuempresa.xclave
                         End With
                         acmd.Execute
                         
                         Set acmd = Nothing
                         DoEvents
                                         
                         '**** Actualizamos Saldos de documento pendiente
                         If rsdetat.Fields(7) = g_tipodolar Then
                            If xmone = g_tiposol Then
                                    VGCNx.Execute "Update  vt_cargo Set cargoapeimppag=isnull(cargoapeimppag,0)+" & CDbl(rsdetat.Fields(8) * xtcam) & "," & _
                                               " cargoapenumpag='" & xnumpag + 1 & "'" & _
                                               " Where documentocargo='" & rsdetat.Fields(4) & "' and cargonumdoc='" & Trim(rsdetat.Fields(6)) & "'" & _
                                               " clientecodigo='" & Trim(Ctr_Ayuda2.xclave) & "'"
                            Else
                                     VGCNx.Execute "Update vt_cargo Set cargoapeimppag=isnull(cargoapeimppag,0)+" & CDbl(rsdetat.Fields(8)) & "," & _
                                                " cargoapenumpag='" & xnumpag + 1 & "'" & _
                                                " Where documentocargo='" & rsdetat.Fields(4) & "' and cargonumdoc='" & Trim(rsdetat.Fields(6)) & "'" & _
                                                " clientecodigo='" & Trim(Ctr_Ayuda2.xclave) & "'"
                                
                            End If
                         ElseIf rsdetat.Fields(7) = g_tiposol Then
                            If xmone = g_tipodolar Then
                                VGCNx.Execute "Update  vt_cargo Set cargoapeimppag=isnull(cargoapeimppag,0)+" & CDbl(rsdetat.Fields(8) / xtcam) & "," & _
                                           " cargoapenumpag='" & xnumpag + 1 & "'" & _
                                           " Where documentocargo='" & rsdetat.Fields(4) & "' and cargonumdoc='" & Trim(rsdetat.Fields(6)) & "'" & _
                                           " clientecodigo='" & Trim(Ctr_Ayuda2.xclave) & "'"
                            Else
                                VGCNx.Execute "Update  vt_cargo Set cargoapeimppag=isnull(cargoapeimppag,0)+" & CDbl(rsdetat.Fields(8)) & "," & _
                                           " cargoapenumpag='" & xnumpag + 1 & "'" & _
                                           " Where documentocargo='" & rsdetat.Fields(4) & "' and cargonumdoc='" & Trim(rsdetat.Fields(6)) & "'" & _
                                           "  and clientecodigo='" & Trim(Ctr_Ayuda2.xclave) & "'"
                            End If
                         End If
                         
                         VGCNx.Execute "Update  vt_cargo " & _
                                     " Set cargoapeflgcan= CASE isnull(cargoapeimpape,0)-isnull(cargoapeimppag,0) WHEN 0 THEN '1' ELSE '0' END ," & _
                                     "   cargoapefeccan='" & rsdetat.Fields(9) & "'" & _
                                     " Where documentocargo='" & rsdetat.Fields(4) & "' and cargonumdoc='" & Trim(rsdetat.Fields(6)) & "'"
                         
                         '**** Actualizamos Saldos del cliente
                         If rsdetat.Fields(7) = g_tipodolar Then
                               VGCNx.Execute "Update  vt_cliente Set clientesaldodolares=isnull(clientesaldodolares,0)-" & CDbl(rsdetat.Fields(8)) & _
                                           " Where clientecodigo='" & Ctr_Ayuda2.xclave & "'"
                         ElseIf rsdetat.Fields(7) = g_tiposol Then
                               VGCNx.Execute "Update  vt_cliente Set clientesaldosoles=isnull(clientesaldosoles,0)-" & CDbl(rsdetat.Fields(8)) & _
                                           " Where clientecodigo='" & Ctr_Ayuda2.xclave & "'"
                         End If
             
                  End If
             End If
             
             ' Registramos datos en Tesoreria
             Set acmd.ActiveConnection = VGGeneral
             acmd.CommandType = adCmdStoredProc
             acmd.CommandText = "te_abonadetalledocumento_pro"
             acmd.CommandTimeout = 0
             acmd.Prepared = True
             With acmd
                 .Parameters("@base") = VGCNx.DefaultDatabase
                 .Parameters("@tipo") = "1"
                 .Parameters("@numrecibo") = Text1(0)
                 .Parameters("@estadoreg") = ""
                 .Parameters("@item") = rsdetat.Fields(0)
                 .Parameters("@emisioncheque") = ""  'IIf(Len(Trim(Text1(2))) = 0, "B", "") ' ver si es cheque
                 .Parameters("@tipodocconcepto") = rsdetat.Fields(1)
                 .Parameters("@numdocumento") = rsdetat.Fields(2)
                 .Parameters("@carabo") = xabono
                 .Parameters("@formacan") = rsdetat.Fields(3)
                 .Parameters("@tdqc") = rsdetat.Fields(4)
                 .Parameters("@ndqc") = Trim(rsdetat.Fields(6))
                 .Parameters("@tipocajabanco") = IIf(Len(Trim(Text1(2))) = 0, "B", "C")
                 .Parameters("@cajabanco") = IIf(Len(Trim(Text1(2))) = 0, Escadena(rsdetat.Fields(5)), Trim(Text1(2)))
                 .Parameters("@numctacte") = Escadena(rsdetat.Fields(10))    'numero de cuenta corriente con tamaño 30
                 .Parameters("@adicionactacte") = "C"
                 .Parameters("@monedadocumento") = xmone
                 .Parameters("@monedacancela") = adll.ComboDato(Combo2)
                 .Parameters("@importesoles") = CDbl(IIf(rsdetat.Fields(7) = g_tiposol, rsdetat.Fields(8), (rsdetat.Fields(8) * xtcam)))
                 .Parameters("@importedolares") = CDbl(IIf(rsdetat.Fields(7) = g_tiposol, (rsdetat.Fields(8) / xtcam), rsdetat.Fields(8)))
          '       .Parameters("@contabledisponi") = Escadena(VGParametros.saldocontadispo)      'sale de empresas
                 .Parameters("@fechacancela") = rsdetat.Fields(9)
                 .Parameters("@observacion") = Escadena(rsdetat.Fields(11))
                 .Parameters("@usuario") = VGUsuario
                 .Parameters("@fechaact") = Now
             End With
             acmd.Execute
             Set acmd = Nothing
             DoEvents
           End If
             rsdetat.MoveNext
         Loop
    End If
    rsdetat.Close
    Set rsdetat = Nothing
    
VGCNx.CommitTrans
    GrabarData = 1
    MsgBox "Los datos han sido grabados satisfactoriamente...!!!", vbInformation, MsgTitle
    Exit Function
error:
  VGCNx.RollbackTrans
  MsgBox "No se pudo Grabar " & Err.Description & " - " & Err.Number, vbInformation, Caption
Exit Function
Resume
End Function

Sub ActualizarLetraCtasCobrar()
  Dim acmd As New ADODB.Command
  
   Set acmd = New ADODB.Command
   Set acmd.ActiveConnection = VGGeneral
   acmd.CommandText = "cc_ingresavarios_pro"
   acmd.CommandType = adCmdStoredProc
   acmd.Prepared = True
  
'   With acmd
'     .Parameters("@base") = VGcnx.DefaultDatabase
'     .Parameters("@tipo") = "1"
'     .Parameters("@tabla") = "vt_cargo"
'     .Parameters("@tipodocu") = Escadena(rsdetav.Fields("td"))
'     .Parameters("@numero") = Escadena(rsdetav.Fields("serie") & rsdetav.Fields("numero"))
'     .Parameters("@cliente") = Escadena(Trim(rsdetav.Fields("cliente")))
'     .Parameters("@vendedor") = Escadena(Ctr_Ayuda2.xclave)
'     .Parameters("@zona") = "01"
'     .Parameters("@apefecemi") = rsdetav.Fields("femision")
'     .Parameters("@moneda") = rsdetav.Fields("moneda")
'     .Parameters("@apeimppag") = CDbl(rsdetav.Fields("importe"))
'     .Parameters("@usuario") = VGusuario
'     .Parameters("@tipocambio") = 0
'     .Parameters("@fechaact") = Date
'     .Parameters("@flagcancel") = 0
'     .Parameters("@tipoplanilla") = Ctr_Ayuda1.xclave
'     .Parameters("@planilla") = xnumplan
'     .Parameters("@vencimiento") = rsdetav.Fields("FVencimiento")
'     .Parameters("@fechaplani") = MBox1.Text
'     .Parameters("@banco") = rsdetav.Fields("banco")
'     .Parameters("@cargoabono") = xcargo
'   End With
   acmd.Execute
   Set acmd = Nothing
   DoEvents





End Sub

Private Sub cmdBotones_Click(Index As Integer)
  Dim nvalor As String
  
  'On Error Resume Next
  
  Select Case Index
    Case 4
       Frame4.Enabled = True
       Call Limpiartexto(Text2, 0, 8)
       Frame4.Enabled = False
       Frame2.Enabled = True
       Call Limpiartexto(Text1, 0, 3)
       Set rsdetat = Nothing
       Call cargar_grilla
       Call ConfigGrid
       Combo1.SetFocus
       Label2(0).Caption = Empty
'       Label2(1).Caption = Empty

    Case 5
       If ValidarGrabacion() = 1 Then
          cmdBotones(5).Enabled = False
          'Grabamos Cabecera de Tesoreria
          If GrabarData() = 1 Then
               
            'Generando Asiento Contable en Linea para cuentas por cobrar
            If VGParametros.sistemaasientoenlinea Then
               Call GeneraAsientoEnlineaTesor(CDate(MBox1.Text), Ctr_Ayuempresa.xclave, "X", Escadena(Text1(0)), 1, "''''", adll.ComboDato(Combo2), IIf(Len(Trim(Text1(2))) = 0, "B", "C"), adll.ComboDato(Combo1))
            End If
            If MsgBox("Desea Imprimir el Recibo ", vbQuestion + vbOKCancel) = vbOK Then
              Call ImprimirRecibo(Escadena(Text1(0)))
            End If
          Else
            MsgBox "No se guardaron los datos....!!!", vbInformation, MsgTitle
          End If
          cmdBotones(5).Enabled = True
          Frame2.Enabled = True
          Call Limpiartexto(Text1, 0, 3)
          Combo1.SetFocus
          Call cmdBotones_Click(4)
       End If
         
    Case 6
      If rsdetat.RecordCount > 0 Then
       nvalor = TDBGrid1.Columns(0).Text
       If rsdetat.RecordCount > 0 Then
          rsdetat.MoveFirst
          Do Until rsdetat.EOF
            If rsdetat.Fields(0) = nvalor Then
              rsdetat.Delete adAffectCurrent
              rsdetat.Update
              Exit Do
            End If
            rsdetat.MoveNext
          Loop
       End If
      End If
      TDBGrid1.Refresh
      Call Totales
      'Call ConfigGrid
      
    Case 7
      Unload Me
  End Select
End Sub

Function ValidarGrabacion() As Integer
ValidarGrabacion = 0
   If rsdetat.RecordCount <= 0 Then
     MsgBox "Falta añadir el Detalle a la Ventana del Browse", vbInformation, Caption
     Exit Function
   End If
  If VGParametros.sistemamultiempresas Then
     If Ctr_Ayuempresa.xclave = "" Then
        MsgBox "Debe ingresar codigo de empresa ", vbInformation
        Exit Function
     End If
   End If
   ValidarGrabacion = 1
End Function

Private Sub Combo1_Click()
  Dim rs As New ADODB.Recordset
  
  Set rs = VGCNx.Execute("select * from te_parametroempresa where empresacodigo='" & VGCodEmpresa & "'")
  If rs.RecordCount > 0 Then
    If adll.ComboDato(Combo1.Text) = "I" Then
        Text1(0) = Right("0000000000" & Trim(CStr(CDbl(IIf(IsNull(rs!empresanumeingreso + 1) Or Len(Trim(rs!empresanumeingreso)) = 0, 1, rs!empresanumeingreso + 1)))), 6)
    ElseIf adll.ComboDato(Combo1.Text) = "E" Then
        Text1(0) = Right("0000000000" & Trim(CStr(CDbl(IIf(IsNull(rs!empresanumegreso + 1) Or Len(Trim(rs!empresanumegreso)) = 0, 1, rs!empresanumegreso + 1)))), 6)
    End If
  End If
  rs.Close
  Set rs = Nothing
 
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    Call Seguir(Combo1, KeyAscii)
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
    Call Seguir(Combo2, KeyAscii)
End Sub

Private Sub Ctr_AyudaCaja_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
Text1(2).Text = Ctr_AyudaCaja.xclave
End Sub

Private Sub Form_Load()
   MostrarForm Me, "C"
   
   Combo1.Clear
   Combo1.AddItem "I- INGRESOS"
   Combo1.AddItem "E- EGRESOS"
   Combo1.ListIndex = 0
   
   Call Ctr_Ayuda2.conexion(VGCNx)
   Call Ctr_AyudaCaja.conexion(VGCNx)
   SQL = " isnull(CajaCuentaxRendir,0)=" & m_docxrendir & " and isnull(Cajafondofijo,0)=" & m_fondofijo
   If VGParametros.listacajas <> "" Then SQL = SQL & " and cajacodigo in (" & VGParametros.listacajas & ")"
   Ctr_AyudaCaja.filtro = SQL
   Call Ctr_Ayuempresa.conexion(VGCNx)
   If VGParametros.sistemamultiempresas Then
      Ctr_Ayuempresa.Visible = True
    Else
      Ctr_Ayuempresa.xclave = VGParametros.empresacodigo
      Ctr_Ayuempresa.Visible = False
      Lblempresa.Visible = False
   End If

    
   Text1(0).Enabled = False
   Call adll.llenacombo(Combo2, "select monedacodigo,monedadescripcion from gr_moneda", VGCNx)
   Combo2.ListIndex = 0
   
   Frame4.Enabled = False
   
   MBox1 = Format(VGParamSistem.fechatrabajo, "dd/mm/yyyy")
   Text1(3) = DatoTipoCambio(VGcnxCT, MBox1.Text)
   Call cargar_grilla
   Call ConfigGrid
   
End Sub

Private Sub MBox1_KeyDown(KeyCode As Integer, Shift As Integer)
  Call Seguir(MBox1, KeyCode)
End Sub

Private Sub MBox1_LostFocus()
 If IsDate(MBox1.Text) Then Text1(3).Text = DatoTipoCambio(VGcnxCT, MBox1.Text)
End Sub

Private Sub MBox2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
       If Format(MBox2.Text, "dd/mm/yyyy") <> Format(MBox1.Text, "dd/mm/yyyy") Then
           MsgBox "La Fecha de Cancelación debe ser la misma para todos los Documentos", vbInformation, Caption
           MBox2.Text = Format(MBox1.Text, "dd/mm/yyyy")
           MBox2.SetFocus
           Exit Sub
       End If

       If Len(Trim(Text1(2))) = 0 Then
           SendKeys "{tab}"
       Else
          Call grabacion
        End If
    End If
End Sub

Public Sub grabacion()
   Dim rb As New ADODB.Recordset
    If Not Text2(3) Like "[TP]" Then
      MsgBox "Solo debe ingresar P ó T", vbInformation, MsgTitle
      Text2(3).SetFocus
      Exit Sub
    End If
    
    Set rb = VGCNx.Execute("select * from cc_tipodocumento where tdocumentotipo='A' and tdocumentoingcobra='1' and tdocumentocodigo='" & Text2(4) & "'")
    If rb.RecordCount = 0 Then
      MsgBox "No existe tipo de documento...Verifique!!", vbInformation, MsgTitle
      rb.Close
      Set rb = Nothing
      Text2(4).SetFocus
      Exit Sub
    End If
    rb.Close
    Set rb = Nothing
    
'    If Text2(4).Text <> "10" Then
'      Set rb = VGCNx.Execute("select * from gr_banco where bancocodigo='" & Text2(5) & "'")
'      If rb.RecordCount = 0 Then
'        MsgBox "No existe el banco indicado .... Verifique!!", vbInformation, MsgTitle
'        rb.Close
'        Set rb = Nothing
'        Text2(5).SetFocus
'        Exit Sub
'      End If
'      rb.Close
'      Set rb = Nothing
'    End If

    Set rb = VGCNx.Execute("select * from gr_moneda where monedacodigo='" & Text2(7) & "'")
    If rb.RecordCount = 0 Then
      MsgBox "No existe moneda .... Verifique!!", vbInformation, MsgTitle
      rb.Close
      Set rb = Nothing
      Text2(7).SetFocus
      Exit Sub
    End If
    rb.Close
    Set rb = Nothing
    Text2(8) = numero(Text2(8))
    
    Set rb = VGCNx.Execute("select * from cc_tipodocumento where tdocumentocodigo='" & Text2(1) & "'")
    If rb.RecordCount > 0 And rb!tdocumentotipo = "C" Then
        Set rb = Nothing
        SQL = "select * from vt_cargo where empresacodigo='" & Ctr_Ayuempresa.xclave & "' And clientecodigo='" & Ctr_Ayuda2.xclave & "'" & _
        " and documentocargo='" & Text2(1) & "' and cargonumdoc='" & Text2(2).Text & "'"
        Set rb = VGCNx.Execute(SQL)
        If rb.RecordCount = 0 Then
            MsgBox "No existe el Nº Documento Referenciado...Verifique!!", vbInformation, MsgTitle
            Text2(2).SetFocus
            Exit Sub
        End If
        Set rb = Nothing
    ElseIf rb.RecordCount > 0 And rb!tdocumentotipo = "A" Then
        SQL = "select * from vt_cargo where empresacodigo='" & Ctr_Ayuempresa.xclave & "' " & _
        " and documentocargo='" & Text2(1) & "' and cargonumdoc='" & Text2(2).Text & "'"
        Set rb = VGCNx.Execute(SQL)
        If rb.RecordCount > 0 Then
            MsgBox "Ya existe el Nº Documento Referenciado...Verifique!!", vbInformation, MsgTitle
            Text2(2).SetFocus
            Exit Sub
        End If
        Set rb = Nothing
    End If
    
    rsdetat.AddNew
    rsdetat.Fields(0) = Escadena(Text2(0))
    rsdetat.Fields(1) = Escadena(Text2(1))
    
    rsdetat.Fields(2) = Escadena(Text2(2))
    
    rsdetat.Fields(3) = Escadena(Text2(3))
    rsdetat.Fields(4) = Escadena(Text2(4))
    rsdetat.Fields(5) = Escadena(Text2(5))
    rsdetat.Fields(6) = Escadena(Text2(6))
    rsdetat.Fields(7) = Escadena(Text2(7))
    rsdetat.Fields(8) = numero(Text2(8))
    rsdetat.Fields(9) = Format(MBox2, "dd/mm/yyyy")
    rsdetat.Fields(10) = Escadena(Text2(9).Text)
    rsdetat.Fields(11) = Escadena(Text2(10).Text)
    rsdetat.Update
    TDBGrid1.Refresh
    Call ConfigGrid
    Call Limpiartexto(Text2, 0, 8)
    MBox2 = Format(MBox1, "dd/mm/yyyy")
    Text2(0) = CStr(CDbl(rsdetat.Fields(0)) + 1)
    Call Totales
    Text2(1).SetFocus
End Sub

Private Sub MBox2_LostFocus()
Call MBox2_KeyDown(13, 0)
End Sub

Private Sub Text1_GotFocus(Index As Integer)
     Call adll.Enfoquetexto(Text1(Index))
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
  Dim rb As New ADODB.Recordset
  On Error Resume Next
  
  If KeyAscii = 13 Then
     If Index = 1 Then
         Set rb = VGCNx.Execute("select * from te_operaciongeneral where operacioncodigo='" & Escadena(Text1(1)) & "' and operacioncontrolaclienteprov='" & IIf(adll.ComboDato(Combo1) = "I", "C", "C") & "'")
         If rb.RecordCount > 0 Then
            Text1(1) = Escadena(rb!operacioncodigo)
            Label2(0) = Escadena(rb!operaciondescripcion)
            If Escadena(rb!operacionvalidacajabancos) = "B" Then
                Ctr_AyudaCaja.Visible = False
'                Text1(2).Enabled = True
                cayuda(1).Enabled = True
                Text1(2) = "": Label2(1) = ""
                Text1(2).Enabled = False
                cayuda(1).Enabled = False
                rb.Close
                Set rb = Nothing
                Combo2.SetFocus
                Text2(6).Enabled = True
                cayuda(5).Enabled = True
                Text2(5).Enabled = True
                Text2(9).Enabled = True
                cayuda(7).Enabled = True
                Exit Sub
            Else
'                Text1(2).Enabled = True
                cayuda(1).Enabled = True
                Ctr_AyudaCaja.Visible = True
                Text2(6).Enabled = False
                cayuda(5).Enabled = False
                Text2(5).Enabled = False
                Text2(9).Enabled = False
                cayuda(7).Enabled = False
                Text1(2).Text = Ctr_AyudaCaja.xclave
'                Text1(2).SetFocus
                Ctr_AyudaCaja.SetFocus
                Set rb = Nothing
                Exit Sub
            End If
         Else
'            Text1(2).Enabled = True
            cayuda(1).Enabled = True
            Text1(1) = "": Label2(0) = "": Text1(2) = "": Label2(1) = ""
         End If
         rb.Close
         Set rb = Nothing
     ElseIf Index = 2 Then
        Set rb = VGCNx.Execute("select * from te_codigocaja where cajacodigo='" & Text1(2) & "'")
        If rb.RecordCount > 0 Then
            Text1(2) = Escadena(rb!cajacodigo)
            Label2(1) = Escadena(rb!cajadescripcion)
        Else
            Text1(2) = ""
            Label2(1) = ""
        End If
        rb.Close
        Set rb = Nothing
     ElseIf Index = 3 Then
        Call Totales
        
        If Not IsDate(MBox1) Then
            MsgBox "Fecha no valida...Verifique!!", vbInformation, MsgTitle
            MBox1.SetFocus
            Exit Sub
        End If
        If Len(Trim(Text1(1))) = 0 Then
            MsgBox "Falta Ingresar Tipo de Operacion...Verifique!!", vbInformation, MsgTitle
            Text1(1).SetFocus
            Exit Sub
        End If
        If Len(Trim(Ctr_Ayuda2.xclave)) = 0 Then
            MsgBox "El Cliente no existe...Verifique!!", vbInformation, MsgTitle
            Ctr_Ayuda2.SetFocus
            Exit Sub
        End If
        If Len(Trim(Text1(3))) = 0 Then
            MsgBox "Falta Ingresar Tipo de Cambio..Verifique!!", vbInformation, MsgTitle
            Text1(3).SetFocus
            Exit Sub
        End If
        Frame4.Enabled = True
        Call Limpiartexto(Text2, 0, 8)
        MBox2 = Format(MBox1, "dd/mm/yyyy")
        If rsdetat.RecordCount = 0 Then
          Text2(0) = 1
        Else
          rsdetat.MoveLast
          Text2(0) = CStr(CDbl(rsdetat.Fields(0)) + 1)
        End If
        Frame2.Enabled = False
        
        Text2(1).SetFocus
        Exit Sub
     End If
     Call Seguir(Text1(Index), 13)
  End If
End Sub

Private Sub Text2_Change(Index As Integer)
  If Index = 3 Then
     Text2(3).Text = UCase(Text2(3).Text)
  End If
End Sub

Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
 Dim rb As New ADODB.Recordset
 
  If KeyAscii = 13 Then
    Text2(Index) = UCase(Text2(Index))
    If Index = 1 Then
        Set rb = VGCNx.Execute("select * from cc_tipodocumento where tdocumentoingplan='1' and tdocumentocodigo='" & Text2(1) & "'")
        If rb.RecordCount = 0 Then
          MsgBox "No existe tipo de documento...Verifique!!", vbInformation, MsgTitle
          rb.Close
          Set rb = Nothing
          Exit Sub
        ElseIf rb!tdocumentonumeauto = 1 Then
'              Text2(11).Text = Left(rb!tdocumentonumerador, 3)
              Text2(2).Text = rb!tdocumentonumerador 'Mid(rb!tdocumentonumerador, 4, Len(rb!tdocumentonumerador) - 3)
        End If
        rb.Close
        Set rb = Nothing
    ElseIf Index = 2 Then
        Set rb = VGCNx.Execute("select * from vt_cargo where clientecodigo='" & Ctr_Ayuda2.xclave & "' and documentocargo='" & Text2(1) & "' and cargonumdoc='" & Text2(2) & "'")
        If rb.RecordCount = 0 Then
          Set rb = VGCNx.Execute("select * from cc_tipodocumento where tdocumentoingplan='1' and tdocumentocodigo='" & Text2(1) & "'")
          If rb!tdocumentotipo = "C" Then
             MsgBox "No existe el Documento...Verifique!!", vbInformation, MsgTitle
             rb.Close
             Set rb = Nothing
             Exit Sub
          End If
        Else
          Set rb = VGCNx.Execute("select * from cc_tipodocumento where tdocumentoingplan='1' and tdocumentocodigo='" & Text2(1) & "'")
          If rb!tdocuemntotipo = "A" Then
             MsgBox "Documento ya fue ingresado ...Verifique!!", vbInformation, MsgTitle
             rb.Close
             Set rb = Nothing
             Exit Sub
          End If
        End If
        rb.Close
        Set rb = Nothing
    ElseIf Index = 3 Then
      Text2(Index) = UCase(Text2(Index))
      If Not Text2(3) Like "[TP]" Then
        MsgBox "Solo debe ingresar P ó T", vbInformation, MsgTitle
        Exit Sub
      End If
    ElseIf Index = 4 Then   'Tipo de cancelacion
      Set rb = VGCNx.Execute("select * from cc_tipodocumento where tdocumentotipo='A' and tdocumentoingcobra='1' and tdocumentocodigo='" & Text2(4) & "'")
      If rb.RecordCount = 0 Then
        MsgBox "No existe tipo de documento...Verifique!!", vbInformation, MsgTitle
        rb.Close
        Set rb = Nothing
        Exit Sub
      End If
      rb.Close
      Set rb = Nothing
    ElseIf Index = 5 Then
      Set rb = VGCNx.Execute("select * from gr_banco where bancocodigo='" & Text2(5) & "'")
      If rb.RecordCount = 0 Then
        MsgBox "No existe el banco indicado .... Verifique!!", vbInformation, MsgTitle
        rb.Close
        Set rb = Nothing
        Exit Sub
      End If
      rb.Close
      Set rb = Nothing
    ElseIf Index = 7 Then
      Set rb = VGCNx.Execute("select * from gr_moneda where monedacodigo='" & Text2(7) & "'")
      If rb.RecordCount = 0 Then
        MsgBox "No existe moneda .... Verifique!!", vbInformation, MsgTitle
        rb.Close
        Set rb = Nothing
        Exit Sub
      End If
      rb.Close
      Set rb = Nothing
    ElseIf Index = 8 Then
       Text2(8) = numero(Text2(8))
       If Text2(8) < 0 Then
        MsgBox "El importe debe ser mayor que cero. Se corregirá el importe", vbInformation, "Aviso"
        Text2(8) = numero(Text2(8) * (-1))
       End If
    ElseIf Index = 9 Then
      Set rb = VGCNx.Execute("select * from te_cuentabancos inner join gr_banco on te_cuentabancos.cbanco_codigo=gr_banco.bancocodigo where gr_banco.bancocodigo='" & Escadena(Text2(5)) & "' and te_cuentabancos.monedacodigo='" & Text2(7) & "' and te_cuentabancos.cbanco_numero='" & Trim(Text2(9)) & "'")
      If rb.RecordCount = 0 Then
        MsgBox "No existe la cuenta corriente del banco indicado .... Verifique!!", vbInformation, MsgTitle
        rb.Close
        Set rb = Nothing
        Exit Sub
      End If
      rb.Close
      Set rb = Nothing
    ElseIf Index = 10 Then
      If Len(Trim(Text1(2))) = 0 Then
          Call grabacion
          Exit Sub
      End If
    End If
    Call Seguir(Text2(Index), KeyAscii)
  End If
End Sub


Private Sub Text2_LostFocus(Index As Integer)
Dim SQL As String
Dim rs As New ADODB.Recordset
      
   If Index = 4 And Trim(Text2(4).Text) = "10" Then
      Text2(6).Enabled = False
      cayuda(5).Enabled = False
      Text2(5).Enabled = False
   Else
      Text2(6).Enabled = True
'      cayuda(5).Enabled = True
'      Text2(5).Enabled = True
   End If
   
   Set rs = New ADODB.Recordset
   If Index = 8 Then
      SQL = "select monedacodigo,isnull(cargoapeimpape,0)-isnull(cargoapeimppag,0) from vt_cargo "
      SQL = SQL & "where empresacodigo='" & Ctr_Ayuempresa.xclave & "' and documentocargo='" & Text2(1).Text & "' and "
      SQL = SQL & "cargonumdoc='" & Trim(Text2(2).Text) & "' and clientecodigo='" & Trim(Ctr_Ayuda2.xclave) & "'"
      Set rs = VGCNx.Execute(SQL)
      If Not rs.BOF And Not rs.EOF Then
        If Text2(7).Text = rs(0) Then
          If Round(numero(Text2(8).Text), 2) > Round(rs(1), 2) Then
            MsgBox "El Monto a Pagar es mayor que el Saldo del Documento", vbInformation, Caption
            Text2(8).SetFocus
            SendKeys "{Home}+{End}"
          End If
        Else
          If rs(0) = g_tiposol Then
            If Round(numero(Text2(8).Text) * MontoCero(Text1(3).Text), 2) > Round(rs(1), 2) Then
              MsgBox "El Monto a Pagar es mayor que el Saldo del Documento", vbInformation, Caption
              Text2(8).SetFocus
              SendKeys "{Home}+{End}"
            End If
          Else
            If Round(numero(Text2(8).Text) / MontoCero(Text1(3).Text), 2) > Round(rs(1), 2) Then
               MsgBox "El Monto a Pagar es mayor que el Saldo del Documento", vbInformation, Caption
               Text2(8).SetFocus
               SendKeys "{Home}+{End}"
            End If
          End If
        End If
      End If
      Set rs = Nothing
   End If
End Sub

Public Function Totales()
    Dim sumas, sumad As Double
    Dim Tsumas, Tsumad As Double
    
    sumas = 0: sumad = 0: Tsumas = 0: Tsumad = 0
    If rsdetat.RecordCount > 0 Then
        rsdetat.MoveFirst
        Do Until rsdetat.EOF
           If rsdetat.Fields(7) = g_tipodolar Then
               sumad = sumad + CDbl(rsdetat.Fields(8))
           ElseIf rsdetat.Fields(7) = g_tiposol Then
               sumas = sumas + CDbl(rsdetat.Fields(8))
           End If
           rsdetat.MoveNext
        Loop
    End If
    If Text1(3) = 0 Or Len(Trim(Text1(3))) = 0 Then Text1(3) = numero(1)
    Tsumad = sumad + (sumas / CDbl(Text1(3)))
    Tsumas = sumad * CDbl(Text1(3)) + sumas

    Label5(0) = numero(Tsumas): Label5(1) = numero(Tsumad)
        
End Function
