VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form FrmProducto 
   Caption         =   "Mantenimiento de Productos"
   ClientHeight    =   7500
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   7500
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmbotones 
      Height          =   1050
      Left            =   3240
      TabIndex        =   31
      Top             =   6360
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
         Picture         =   "FrmProducto.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   36
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
         Picture         =   "FrmProducto.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   35
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
         Left            =   2440
         Picture         =   "FrmProducto.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   34
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
         Picture         =   "FrmProducto.frx":0CC6
         Style           =   1  'Graphical
         TabIndex        =   33
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
         Picture         =   "FrmProducto.frx":1108
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   180
         Width           =   915
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6135
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   10821
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Consulta"
      TabPicture(0)   =   "FrmProducto.frx":154A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "TDBGridProducto"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Mantenimiento"
      TabPicture(1)   =   "FrmProducto.frx":1566
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cCancela"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cAcepta"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
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
         Left            =   4440
         TabIndex        =   13
         Top             =   5400
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
         Left            =   6360
         TabIndex        =   14
         Top             =   5400
         Width           =   1335
      End
      Begin VB.Frame Frame1 
         Height          =   4875
         Left            =   360
         TabIndex        =   17
         Top             =   375
         Width           =   11385
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
            Index           =   11
            Left            =   7560
            MaxLength       =   8
            TabIndex        =   12
            Top             =   4455
            Width           =   2145
         End
         Begin VB.ComboBox cmbMoneda 
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
            Left            =   7560
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   3240
            Width           =   3375
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
            Index           =   10
            Left            =   7560
            TabIndex        =   10
            Top             =   3675
            Width           =   2145
         End
         Begin VB.CheckBox chk 
            Height          =   375
            Index           =   0
            Left            =   2400
            TabIndex        =   2
            Top             =   1560
            Width           =   615
         End
         Begin VB.ComboBox cmbUnidad 
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
            Left            =   2400
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   960
            Width           =   3135
         End
         Begin VB.ComboBox cmbGrupoVta 
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
            Left            =   7560
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   1320
            Width           =   3375
         End
         Begin VB.Frame Frame2 
            Caption         =   "Unidad Medida Referencial"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2175
            Left            =   240
            TabIndex        =   27
            Top             =   2160
            Width           =   5295
            Begin VB.TextBox txt 
               Enabled         =   0   'False
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
               Index           =   9
               Left            =   2160
               TabIndex        =   39
               Top             =   1680
               Width           =   1305
            End
            Begin VB.TextBox txt 
               Enabled         =   0   'False
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
               Left            =   2160
               TabIndex        =   37
               Top             =   1080
               Width           =   2985
            End
            Begin VB.TextBox txt 
               Enabled         =   0   'False
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
               Left            =   2160
               TabIndex        =   30
               Top             =   480
               Width           =   1065
            End
            Begin VB.Label lbl 
               Caption         =   "Código"
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
               Left            =   240
               TabIndex        =   38
               Top             =   600
               Width           =   795
            End
            Begin VB.Label lbl 
               Caption         =   "Factor Conversión"
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
               Left            =   240
               TabIndex        =   29
               Top             =   1680
               Width           =   1875
            End
            Begin VB.Label lbl 
               Caption         =   "Descripción"
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
               Left            =   240
               TabIndex        =   28
               Top             =   1080
               Width           =   1275
            End
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
            Left            =   7560
            MaxLength       =   1
            TabIndex        =   8
            Top             =   2760
            Width           =   3345
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
            Index           =   4
            Left            =   7560
            MaxLength       =   3
            TabIndex        =   7
            Top             =   2280
            Width           =   3345
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
            Left            =   7560
            MaxLength       =   30
            TabIndex        =   4
            Top             =   825
            Width           =   3345
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
            Left            =   7560
            MaxLength       =   80
            TabIndex        =   3
            Top             =   340
            Width           =   3345
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
            Left            =   2400
            MaxLength       =   8
            TabIndex        =   0
            Top             =   360
            Width           =   1185
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
            Index           =   6
            Left            =   7560
            MaxLength       =   8
            TabIndex        =   11
            Top             =   4065
            Width           =   2145
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
            Index           =   3
            Left            =   7560
            MaxLength       =   3
            TabIndex        =   6
            Top             =   1800
            Width           =   3345
         End
         Begin VB.Label lbl 
            Caption         =   "Almacen"
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
            Left            =   5760
            TabIndex        =   43
            Top             =   4455
            Width           =   1470
         End
         Begin VB.Label lbl 
            Caption         =   "Moneda"
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
            Index           =   15
            Left            =   5760
            TabIndex        =   42
            Top             =   3240
            Width           =   1575
         End
         Begin VB.Label lbl 
            Caption         =   "Precio Venta"
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
            Left            =   5760
            TabIndex        =   41
            Top             =   3720
            Width           =   1575
         End
         Begin VB.Label lbl 
            Caption         =   "Medida Referencia"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   9
            Left            =   240
            TabIndex        =   40
            Top             =   1560
            Width           =   1845
         End
         Begin VB.Label lbl 
            Caption         =   "Unidad Medida"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   8
            Left            =   240
            TabIndex        =   26
            Top             =   960
            Width           =   1605
         End
         Begin VB.Label lbl 
            Caption         =   "Familia"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   7
            Left            =   5760
            TabIndex        =   25
            Top             =   1800
            Width           =   825
         End
         Begin VB.Label lbl 
            Caption         =   "Categoría"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   6
            Left            =   5760
            TabIndex        =   24
            Top             =   2280
            Width           =   1605
         End
         Begin VB.Label lbl 
            Caption         =   "Tipo"
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
            Left            =   5760
            TabIndex        =   23
            Top             =   2760
            Width           =   1635
         End
         Begin VB.Label lbl 
            Caption         =   "Porc. Impuesto"
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
            Left            =   5760
            TabIndex        =   22
            Top             =   4110
            Width           =   1695
         End
         Begin VB.Label lbl 
            Caption         =   "Grupo Venta"
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
            Left            =   5760
            TabIndex        =   21
            Top             =   1350
            Width           =   1335
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
            Height          =   315
            Index           =   2
            Left            =   5760
            TabIndex        =   20
            Top             =   885
            Width           =   1485
         End
         Begin VB.Label lbl 
            Caption         =   "Descripción"
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
            Left            =   5760
            TabIndex        =   19
            Top             =   440
            Width           =   1275
         End
         Begin VB.Label lbl 
            Caption         =   "Código"
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
            Left            =   240
            TabIndex        =   18
            Top             =   360
            Width           =   1080
         End
      End
      Begin TrueOleDBGrid70.TDBGrid TDBGridProducto 
         Height          =   5535
         Left            =   -74880
         TabIndex        =   16
         Top             =   480
         Width           =   11775
         _ExtentX        =   20770
         _ExtentY        =   9763
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
         Splits(0).DividerColor=   13160660
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
         CellTipsWidth   =   0
         DeadAreaBackColor=   13160660
         RowDividerColor =   13160660
         RowSubDividerColor=   13160660
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
Attribute VB_Name = "FrmProducto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim modoinsert, modoedit As Boolean
Dim i_filaorigen, i_indexcombo As Integer
Dim i_valorcodigo As String
Dim adll As New dllgeneral.dll_general
''''''''''''''''''''''''
Dim ArregloUnidades()
Dim ArregloGrupoVtas()
Dim ArregloMoneda()

Private Sub cAcepta_Click()
    
   Dim rs As New ADODB.Recordset
   Dim SQL As String
   Dim J As Integer
   Dim f_porcentaje As Double
   Dim f_factorconversion As Double
   Dim s_codigounidad, s_codigomoneda, s_codigogrupovta As String
   Dim f_precioventa As Double
   
   On Error GoTo CONTROLERRORES
   ''''''''
    If cmbMoneda.ListIndex <> -1 Then
        If Not (Val(txt(10)) > 0) Then
            MsgBox "Ingrese Precio de Venta", vbCritical, "AVISO"
            Exit Sub
        End If
    End If
    
    If txt(10) <> "" Then
        If Not cmbMoneda.ListIndex <> -1 Then
            MsgBox "Ingrese Moneda", vbCritical, "AVISO"
            Exit Sub
        End If
    End If
   
    If chk(0).Value = 1 Then
        If txt(7) = "" Then
            MsgBox "La unidad de medida seleccionada no tiene asociada" & _
            " una unidad referencial", vbInformation, "AVISO"
            chk(0).Value = 0
            fncUnidadMedidaReferencial (0)
            Exit Sub
        End If
    End If
   
   
     If txt(6) = "" Then
         f_porcentaje = 0
     Else
         f_porcentaje = txt(6) / 100
     End If
     If txt(9) = "" Then
         f_factorconversion = 0
     Else
         f_factorconversion = txt(9)
     End If
     If txt(10) = "" Then
         f_precioventa = 0
     Else
         f_precioventa = txt(10)
     End If
   
     If cmbUnidad.ListIndex <> -1 Then
        s_codigounidad = ArregloUnidades(0, cmbUnidad.ListIndex)
     Else
        s_codigounidad = ""
     End If
     If cmbGrupoVta.ListIndex <> -1 Then
        s_codigogrupovta = ArregloGrupoVtas(0, cmbGrupoVta.ListIndex)
     Else
        s_codigogrupovta = ""
     End If
     If cmbMoneda.ListIndex <> -1 Then
        s_codigomoneda = ArregloMoneda(0, cmbMoneda.ListIndex)
     Else
        s_codigomoneda = ""
     End If
        
   
   If modoinsert = True Then
   
   
         If Validar_CodigosDuplicados("INSERT") = True Then
            If adll.VerificaDatoExistente(VGcnx, "select * from vt_producto where productocodigo='" & txt(1) & "' and almacencodigo='" & txt(11) & "'") = 1 Then
              MsgBox "Código ya existe", vbCritical, "Error"
              cAcepta.Enabled = False
              Exit Sub
            End If
          End If
          
         If Validar_DescripcionesDuplicadas("INSERT") = True Then
            MsgBox "Descripción ya existe", vbCritical, "Error"
            cAcepta.Enabled = False
            Exit Sub
          End If
               
          SQL = "INSERT INTO vt_producto " & _
               "(productocodigo,productodescripcion,productodescrcorta," & _
               "grupovtacodigo,productofamiliacodigo,productocategoriacodigo," & _
               "productotipo,unidadcodigo,productoporcimpto, " & _
               "productoestunidreferencia,unidadreferencial," & _
               "unidadfactorconv,usuariocodigo,fechaact,productoprecvta,monedacodigo,almacencodigo" & _
               ") VALUES " & _
               "('" & txt(0) & "','" & txt(1) & "','" & txt(2) & "','" & _
                s_codigogrupovta & _
               "','" & txt(3) & "','" & txt(4) & "','" & txt(5) & "','" & _
                s_codigounidad & _
               "'," & f_porcentaje & "," & chk(0).Value & ",'" & txt(7) & "'," & _
               f_factorconversion & ",'" & g_usuario & "','" & Date & "'," & f_precioventa & ",'" & s_codigomoneda & "','" & txt(11) & "')"

          VGcnx.Execute SQL
          
          'Aumentamos en lista de precios
           SQL = "INSERT INTO listapre1 " & _
               "(productocodigo,productodescripcion,productodescrcorta," & _
               "grupovtacodigo,productofamiliacodigo,productocategoriacodigo," & _
               "productotipo,unidadcodigo,productoporcimpto, " & _
               "productoestunidreferencia,unidadreferencial," & _
               "unidadfactorconv,productoprecvta,monedacodigo,almacencodigo" & _
               ") VALUES " & _
               "('" & txt(0) & "','" & txt(1) & "','" & txt(2) & "','" & _
                s_codigogrupovta & _
               "','" & txt(3) & "','" & txt(4) & "','" & txt(5) & "','" & _
                s_codigounidad & _
               "'," & f_porcentaje & "," & chk(0).Value & ",'" & txt(7) & "'," & _
               f_factorconversion & "," & f_precioventa & ",'" & s_codigomoneda & "','" & txt(11) & "')"
        VGcnx.Execute SQL
        
    ElseIf modoedit = True Then
   
             If Validar_CodigosDuplicados("UPDATE", i_filaorigen) = True Then
               MsgBox "Código ya existe", vbCritical, "Error"
               cAcepta.Enabled = False
               Exit Sub
             End If
             
             If Validar_DescripcionesDuplicadas("UPDATE", i_filaorigen) = True Then
               MsgBox "Descripción ya existe", vbCritical, "Error"
               cAcepta.Enabled = False
               Exit Sub
             End If
                                 
            SQL = "UPDATE vt_producto SET " & _
               "productodescripcion='" & txt(1) & "'," & _
               "productodescrcorta='" & txt(2) & "'," & _
               "grupovtacodigo='" & s_codigogrupovta & "'," & _
               "productofamiliacodigo='" & txt(3) & "'," & _
               "productocategoriacodigo='" & txt(4) & "'," & _
               "productotipo='" & txt(5) & "'," & _
               "unidadcodigo='" & s_codigounidad & "'," & _
               "productoporcimpto=" & f_porcentaje & "," & _
               "productoestunidreferencia=" & chk(0).Value & "," & _
               "unidadreferencial='" & txt(7) & "'," & _
               "unidadfactorconv=" & f_factorconversion & "," & _
               "productoprecvta=" & f_precioventa & ", " & _
               "monedacodigo='" & s_codigomoneda & "'  " & _
               "WHERE productocodigo='" & txt(0) & "'"
    
            VGcnx.Execute SQL
            
            SQL = "UPDATE listapre1 SET " & _
               "productodescripcion='" & txt(1) & "'," & _
               "productodescrcorta='" & txt(2) & "'," & _
               "grupovtacodigo='" & s_codigogrupovta & "'," & _
               "productofamiliacodigo='" & txt(3) & "'," & _
               "productocategoriacodigo='" & txt(4) & "'," & _
               "productotipo='" & txt(5) & "'," & _
               "unidadcodigo='" & s_codigounidad & "'," & _
               "productoporcimpto=" & f_porcentaje & "," & _
               "productoestunidreferencia=" & chk(0).Value & "," & _
               "unidadreferencial='" & txt(7) & "'," & _
               "unidadfactorconv=" & f_factorconversion & "," & _
               "productoprecvta=" & f_precioventa & ", " & _
               "monedacodigo='" & s_codigomoneda & "'  " & _
               "WHERE productocodigo='" & txt(0) & "'"
            VGcnx.Execute SQL

            
  End If
 '******************************************************************************************
        
 TDBGridProducto.Refresh
      
 Mostrar_Data
 MostrarOcultar_Botones (True)
 '''''''''
 modoinsert = False
 modoedit = False
 SSTab1.TabEnabled(0) = True
 '''''''''
 'rs.Close
 'Set rs = Nothing
Exit Sub
CONTROLERRORES:
   If Err Then
       MsgBox Err.Number & "-" & Err.Description, vbInformation, "ERROR"
       Err = 0
       'VGgeneral.RollbackTrans
       Resume Next
    End If
       
End Sub

Private Sub cCancela_Click()
    SSTab1.TabEnabled(0) = True
    SSTab1.Tab = 0
    SSTab1.SetFocus
    MostrarOcultar_Botones (True)
    modoinsert = False
    modoedit = False
End Sub

Private Sub chk_Click(Index As Integer)
     cAcepta.Enabled = Validar_DatosNulos()
     If cmbUnidad.ListIndex <> -1 Then
         Call fncUnidadMedidaReferencial(chk(0).Value)
     Else
        MsgBox "Seleccione Unidad de Medida", vbInformation, "AVISO"
     End If
End Sub

Private Sub chk_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub cmbGrupoVta_Click()
    cAcepta.Enabled = Validar_DatosNulos()
End Sub

Private Sub cmbGrupoVta_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub cmbMoneda_Click()
    cAcepta.Enabled = Validar_DatosNulos()
End Sub

Private Sub cmbMoneda_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub cmbUnidad_Click()
   cAcepta.Enabled = Validar_DatosNulos()
   If i_indexcombo <> cmbUnidad.ListIndex Then
    fncUnidadMedidaReferencial (0)
    chk(0).Value = 0
   End If
End Sub

Private Sub cmbUnidad_DropDown()
   i_indexcombo = cmbUnidad.ListIndex
End Sub

Private Sub cmbUnidad_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub cmdBotones_Click(Index As Integer)
  Dim J As Integer
  Dim SQL As String
  Dim OBJ As Object
  
  On Error GoTo CONTROLERRORES
  
  SSTab1.TabEnabled(1) = True
  '''''
  Select Case Index
     Case 0   'nuevo
            For Each OBJ In Me.Controls
               If TypeOf OBJ Is TextBox Then
                    OBJ.Text = ""
                End If
                If TypeOf OBJ Is ComboBox Then
                    OBJ.ListIndex = -1
                End If
                If TypeOf OBJ Is CheckBox Then
                OBJ.Value = 0
                End If
            Next
            SSTab1.Tab = 1
            modoinsert = True
            MostrarOcultar_Botones (False)
            txt(0).SetFocus
        
     Case 1   'modificar
     
         If TDBGridProducto.Row < 0 Then
            Exit Sub
         End If
         
             Call fncSeleccionaCombo(Trim(TDBGridProducto.Columns(8).Text), cmbUnidad, ArregloUnidades)
             Call fncSeleccionaCombo(Trim(TDBGridProducto.Columns(3).Text), cmbGrupoVta, ArregloGrupoVtas)
             Call fncSeleccionaCombo(Trim(TDBGridProducto.Columns(16).Text), cmbMoneda, ArregloMoneda)
             
             i_valorcodigo = Trim(TDBGridProducto.Columns(0).Text)
             txt(7) = IIf(IsNull(ArregloUnidades(2, cmbUnidad.ListIndex)), "", ArregloUnidades(2, cmbUnidad.ListIndex))
'             txt(8) = IIf(IsNull(ArregloUnidades(3, cmbUnidad.ListIndex)), "", ArregloUnidades(3, cmbUnidad.ListIndex))
'             txt(9) = IIf(IsNull(ArregloUnidades(4, cmbUnidad.ListIndex)), "", ArregloUnidades(4, cmbUnidad.ListIndex))
            
             txt(0) = Trim(TDBGridProducto.Columns(0).Text)
             txt(1) = Trim(TDBGridProducto.Columns(1).Text)
             txt(2) = Trim(TDBGridProducto.Columns(2).Text)
             txt(3) = Trim(TDBGridProducto.Columns(5).Text)
             txt(4) = Trim(TDBGridProducto.Columns(6).Text)
             txt(5) = Trim(TDBGridProducto.Columns(7).Text)
             txt(6) = Trim(TDBGridProducto.Columns(10).Text)
             txt(7) = Trim(TDBGridProducto.Columns(12).Text)
             txt(8) = Trim(TDBGridProducto.Columns(13).Text)
             txt(9) = Trim(TDBGridProducto.Columns(14).Text)
             txt(10) = Trim(TDBGridProducto.Columns(15).Text)
         
            If TDBGridProducto.Columns(11).Value = False Then
                 chk(0).Value = 0
            ElseIf TDBGridProducto.Columns(11).Value = True Then
                 chk(0).Value = 1
            End If
         
                 
        modoedit = True
        SSTab1.Tab = 1
        MostrarOcultar_Botones (False)
        i_filaorigen = TDBGridProducto.Row
        txt(0).SetFocus
      
        '''''''''
      
     Case 2   'eliminar
        If TDBGridProducto.Row < 0 Then
            Exit Sub
        End If
     
       If MsgBox("Desea eliminar el registro?", vbYesNo + vbDefaultButton2, "AVISO") = vbYes Then
          SQL = "DELETE FROM vt_producto WHERE productocodigo = " & TDBGridProducto.Columns(0).Text
          VGcnx.Execute SQL
          
          SQL = "DELETE FROM listapre1 WHERE productocodigo = " & TDBGridProducto.Columns(0).Text
          VGcnx.Execute SQL
          
          Mostrar_Data
       End If
        
     Case 3   'imprimir
         Call Imprimir("MantProducto.rpt")
     Case 4  ' salir
       Unload Me
  End Select
Exit Sub
CONTROLERRORES:
   If Err Then
       MsgBox Err.Number & "-" & Err.Description, vbInformation, "ERROR"
       Err = 0
       'VGgeneral.RollbackTrans
       Resume Next
    End If

End Sub

Private Sub Form_Load()
 MostrarForm Me, "C2"
 Mostrar_Data
 Setear_Controles
 cAcepta.Enabled = False
 SSTab1.TabEnabled(1) = False
End Sub

Public Function Mostrar_Data()
  Dim SQL As String
  Dim rs As New ADODB.Recordset
  Dim I As Integer
    
      SQL = "SELECT a.productocodigo as Código, a.productodescripcion as Descripción," & _
      "a.productodescrcorta as 'Descripcion Corta'," & _
      "a.grupovtacodigo as 'Cód. Grupo Venta' , b.grupovtadescripcion as 'Grupo Venta'," & _
      "a.productofamiliacodigo as 'Familia', a.productocategoriacodigo as 'Categoría'," & _
      "a.productotipo as 'Tipo'," & _
      "a.unidadcodigo as 'Cód.Unid.Med.',c.unidaddescripcion as 'Desc.Unid.Med.'," & _
      "a.productoporcimpto*100 as 'Porc.Impto', a.productoestunidreferencia as 'Medida Refere.'," & _
      "a.unidadreferencial as 'Cód.Un.Med.Refer.',d.unidaddescripcion as 'Des.Un.Med.Refer.'," & _
      "a.unidadfactorconv as 'Factor Conversion'," & _
      "a.productoprecvta as 'Precio Venta'," & _
      "a.monedacodigo, e.monedadescripcion" & _
      " " & _
      "FROM  vt_producto a " & _
      "  LEFT JOIN  vt_grupoventa b ON a.grupovtacodigo = b.grupovtacodigo" & _
      "      LEFT JOIN    vt_unidad c ON a.unidadcodigo = c.unidadcodigo" & _
      "      LEFT OUTER JOIN vt_unidad d ON a.unidadreferencial = d.unidadcodigo " & _
      "      LEFT OUTER JOIN gr_moneda e ON a.monedacodigo = e.monedacodigo " & _
      "ORDER BY a.productocodigo"
      
      Set rs = VGcnx.Execute(SQL)
      Set TDBGridProducto.DataSource = rs
    
     '' COMBO UNIDAD:
      SQL = "SELECT a.unidadcodigo,a.unidaddescripcion,a.unidadreferencial," & _
      "b.unidaddescripcion,a.unidadfactorconv " & _
      "FROM vt_unidad a LEFT OUTER JOIN vt_unidad b " & _
      "ON a.unidadreferencial=b.unidadcodigo " & _
      "WHERE a.estadoreg = 1 ORDER BY a.unidadcodigo "
      Set rs = VGcnx.Execute(SQL)
      If rs.RecordCount > 0 Then
        ReDim ArregloUnidades(0 To 4, 0 To rs.RecordCount - 1)
        Call fncLlenarArreglo_Combo(rs, cmbUnidad, ArregloUnidades, 4)
      End If
      '' COMBO GRUPO VENTA:
      SQL = "SELECT grupovtacodigo,grupovtadescripcion FROM vt_grupoventa " & _
      "ORDER BY grupovtacodigo "
      Set rs = VGcnx.Execute(SQL)
      If rs.RecordCount > 0 Then
        ReDim ArregloGrupoVtas(0 To 1, 0 To rs.RecordCount - 1)
        Call fncLlenarArreglo_Combo(rs, cmbGrupoVta, ArregloGrupoVtas, 1)
      End If
      ' COMBO MONEDA:
      SQL = "SELECT monedacodigo,monedadescripcion " & _
      "FROM gr_moneda " & _
      "ORDER BY monedacodigo "
      Set rs = VGcnx.Execute(SQL)
      If rs.RecordCount > 0 Then
        ReDim ArregloMoneda(0 To 1, 0 To rs.RecordCount - 1)
        Call fncLlenarArreglo_Combo(rs, cmbMoneda, ArregloMoneda, 1)
      End If
      
      Setear_Controles
        
 TDBGridProducto.Refresh
 Set rs = Nothing
 SSTab1.Tab = 0
  
End Function

Private Function Setear_Controles()
Dim I As Integer

    For I = 0 To TDBGridProducto.Columns.Count - 1
        Select Case I
            'Case 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 12, 13, 14
            '    TDBGridProducto.Columns(i).ValueItems.Presentation = dbgNormal
            Case 3, 12
                TDBGridProducto.Columns(I).Visible = False
            Case 11
                TDBGridProducto.Columns(I).ValueItems.Presentation = dbgCheckBox
                TDBGridProducto.Columns(I).Width = 700
            Case 14, 10
                TDBGridProducto.Columns(I).Width = 600
            Case Else
                TDBGridProducto.Columns(I).Width = 850
    End Select
    Next I
    
End Function

Private Function Validar_DatosNulos() As Boolean

Validar_Ingreso = False

                If Trim(txt(0)) <> "" And Trim(txt(1)) <> "" And Trim(cmbGrupoVta.Text) <> "" _
                  And Trim(cmbUnidad.Text) <> "" Then
                    Validar_DatosNulos = True
                    Exit Function
                End If

End Function


Private Sub SSTab1_Click(PreviousTab As Integer)
    SSTab1.TabEnabled(PreviousTab) = False
    cAcepta.Enabled = False
End Sub


Private Sub txt_Change(Index As Integer)
cAcepta.Enabled = Validar_DatosNulos()
End Sub

Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)  ' Salta con Enter
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    cAcepta.Enabled = Validar_DatosNulos()
    
    'Ingresar Mayusculas:
    If KeyAscii > 96 And KeyAscii < 123 Then
        KeyAscii = KeyAscii - 32
    End If

End Sub

Private Function Validar_CodigosDuplicados(operacion As String, Optional ByVal filaorigen As Integer) As Boolean
               
Validar_CodigosDuplicados = False
                        
   TDBGridProducto.MoveFirst
   Do Until TDBGridProducto.EOF
       If operacion = "INSERT" Then
          If Trim(txt(0)) = _
             Trim(TDBGridProducto.Columns.Item(0).Value) Then
                 Validar_CodigosDuplicados = True
                 Exit Function
          End If
       ElseIf operacion = "UPDATE" Then
          If Trim(txt(0)) = _
             Trim(TDBGridProducto.Columns.Item(0).Value) And _
             TDBGridProducto.Row <> filaorigen Then
                 Validar_CodigosDuplicados = True
                 Exit Function
          End If
       End If
       TDBGridProducto.MoveNext
    Loop
               
End Function

Private Function MostrarOcultar_Botones(valor As Boolean)
    frmbotones.Visible = valor
End Function

Private Function fncSeleccionaCombo(ValorCodigo As String, Cbo As ComboBox, Arreglo As Variant)
Dim I As Integer
    For I = 0 To UBound(Arreglo, 2)
       If ValorCodigo = Arreglo(0, I) Then
         Cbo.ListIndex = I
         Exit Function
       End If
    Next I
End Function

Private Function fncLlenarArreglo_Combo(rs As Recordset, Cbo As ComboBox, Arreglo As Variant, dimensiones As Integer)
Dim I As Integer
Dim J As Integer

    I = 0
    Cbo.Clear
    Do Until rs.EOF
        Cbo.AddItem (Trim(rs(1)))
        For J = 0 To dimensiones
            Arreglo(J, I) = Trim(rs(J))
        Next J
        rs.MoveNext
        I = I + 1
    Loop
End Function

Private Function fncUnidadMedidaReferencial(tipo As Integer)
If modoedit = True Then
    If tipo = 0 Then                'checkbox desmarcado
        txt(7) = ""
        txt(8) = ""
        txt(9) = ""
    ElseIf tipo = 1 Then            'checkbox marcado
        txt(7) = ArregloUnidades(2, cmbUnidad.ListIndex) & ""
        txt(8) = ArregloUnidades(3, cmbUnidad.ListIndex) & ""
        txt(9) = ArregloUnidades(4, cmbUnidad.ListIndex) & ""
    End If
ElseIf modoinsert = True Then
   If tipo = 0 Then                'checkbox desmarcado
        txt(7) = ""
        txt(8) = ""
        txt(9) = ""
    ElseIf tipo = 1 Then            'checkbox marcado
        txt(7) = ArregloUnidades(2, cmbUnidad.ListIndex) & ""
        txt(8) = ArregloUnidades(3, cmbUnidad.ListIndex) & ""
        txt(9) = ArregloUnidades(4, cmbUnidad.ListIndex) & ""
    End If
End If
End Function
Public Function Formatear_Codigo(indice As Integer) As String
Dim cadena As String
Dim I As Integer

cadena = ""
For I = 0 To txt(indice).MaxLength
    cadena = cadena & "0"
Next I

txt(indice) = Right(cadena & Trim(txt(indice)), txt(indice).MaxLength)

End Function

Private Sub txt_LostFocus(Index As Integer)
If txt(Index) <> "" Then
    If Index = 0 Then
        'Call Formatear_Codigo(Index)
    End If
    If Index = 6 Or Index = 10 Then
        txt(Index).Text = Format(Val(txt(Index).Text), "#,###,##0.00")
    End If
End If
End Sub

Private Function Validar_DescripcionesDuplicadas(operacion As String, Optional ByVal filaorigen As Integer) As Boolean
               
Validar_DescripcionesDuplicadas = False
                        
     TDBGridProducto.MoveFirst
     Do Until TDBGridProducto.EOF
        If operacion = "INSERT" Then
           If Trim(txt(1)) = _
              Trim(TDBGridProducto.Columns.Item(1).Value) Then
                  Validar_DescripcionesDuplicadas = True
                  Exit Function
           End If
        ElseIf operacion = "UPDATE" Then
           If Trim(txt(1)) = _
              Trim(TDBGridProducto.Columns.Item(1).Value) And _
              TDBGridProducto.Row <> filaorigen Then
                   Validar_DescripcionesDuplicadas = True
                   Exit Function
           End If
       End If
       TDBGridProducto.MoveNext
    Loop
               
End Function

