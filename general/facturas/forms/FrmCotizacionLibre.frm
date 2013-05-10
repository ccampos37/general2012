VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form FrmCotizacionLibre 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cotizacion Libre"
   ClientHeight    =   9150
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12030
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9150
   ScaleWidth      =   12030
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Panel 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      ScaleHeight     =   285
      ScaleWidth      =   11970
      TabIndex        =   50
      Top             =   8805
      Width           =   12030
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8565
      Left            =   0
      TabIndex        =   51
      Top             =   0
      Width           =   11925
      _ExtentX        =   21034
      _ExtentY        =   15108
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Consulta"
      TabPicture(0)   =   "FrmCotizacionLibre.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "LblReg"
      Tab(0).Control(1)=   "Label7"
      Tab(0).Control(2)=   "Label8"
      Tab(0).Control(3)=   "Label9"
      Tab(0).Control(4)=   "DtFechaHasta"
      Tab(0).Control(5)=   "TDBGrid2"
      Tab(0).Control(6)=   "cmdBotones(0)"
      Tab(0).Control(7)=   "cmdBotones(1)"
      Tab(0).Control(8)=   "cmdBotones(2)"
      Tab(0).Control(9)=   "cmdBotones(4)"
      Tab(0).Control(10)=   "TxtNro"
      Tab(0).Control(11)=   "DtFechaDesde"
      Tab(0).Control(12)=   "TxtCliente"
      Tab(0).Control(13)=   "CmdBuscar"
      Tab(0).ControlCount=   14
      TabCaption(1)   =   "Mantenimiento"
      TabPicture(1)   =   "FrmCotizacionLibre.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "SSTab2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "oCrystalReport"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "TDBGrid1"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Frame3"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Fr2(2)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Fr2(1)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "cmdBotones(11)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "cmdBotones(12)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).ControlCount=   8
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Buscar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Left            =   -69510
         Picture         =   "FrmCotizacionLibre.frx":0038
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   585
         Width           =   1140
      End
      Begin VB.TextBox TxtCliente 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -73065
         MaxLength       =   50
         TabIndex        =   1
         Top             =   945
         Width           =   3390
      End
      Begin MSComCtl2.DTPicker DtFechaDesde 
         Height          =   285
         Left            =   -73065
         TabIndex        =   2
         Top             =   1305
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   503
         _Version        =   393216
         Format          =   117047297
         CurrentDate     =   39763
         MaxDate         =   44196
         MinDate         =   36526
      End
      Begin VB.TextBox TxtNro 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -73065
         MaxLength       =   11
         TabIndex        =   0
         Top             =   585
         Width           =   1365
      End
      Begin VB.CommandButton cmdBotones 
         Caption         =   "&Salir"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Index           =   4
         Left            =   -64425
         Style           =   1  'Graphical
         TabIndex        =   131
         Top             =   585
         Width           =   1140
      End
      Begin VB.CommandButton cmdBotones 
         Caption         =   "&Eliminar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Index           =   2
         Left            =   -65595
         Style           =   1  'Graphical
         TabIndex        =   130
         Top             =   585
         Width           =   1095
      End
      Begin VB.CommandButton cmdBotones 
         Caption         =   "E&ditar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Index           =   1
         Left            =   -66810
         Style           =   1  'Graphical
         TabIndex        =   129
         Top             =   585
         Width           =   1140
      End
      Begin VB.CommandButton cmdBotones 
         Caption         =   "&Nuevo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Index           =   0
         Left            =   -68025
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   585
         Width           =   1140
      End
      Begin VB.CommandButton cmdBotones 
         Caption         =   "&Cancelar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1050
         Index           =   12
         Left            =   10485
         Style           =   1  'Graphical
         TabIndex        =   127
         Top             =   7335
         Width           =   1215
      End
      Begin VB.CommandButton cmdBotones 
         Caption         =   "&Acepta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1050
         Index           =   11
         Left            =   9135
         Style           =   1  'Graphical
         TabIndex        =   126
         Top             =   7335
         Width           =   1215
      End
      Begin VB.Frame Fr2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Index           =   1
         Left            =   30
         TabIndex        =   111
         Top             =   2880
         Width           =   11805
         Begin VB.CommandButton cAyuda 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   3
            Left            =   3660
            TabIndex        =   41
            Top             =   420
            Width           =   285
         End
         Begin MSMask.MaskEdBox MBox2 
            Height          =   285
            Index           =   0
            Left            =   1530
            TabIndex        =   31
            Top             =   450
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   503
            _Version        =   393216
            ClipMode        =   1
            Appearance      =   0
            MaxLength       =   6
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MBox2 
            Height          =   285
            Index           =   1
            Left            =   2340
            TabIndex        =   32
            Top             =   450
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   8
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MBox2 
            Height          =   285
            Index           =   2
            Left            =   7710
            TabIndex        =   112
            Top             =   450
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   12648447
            Enabled         =   0   'False
            MaxLength       =   8
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MBox2 
            Height          =   285
            Index           =   3
            Left            =   8610
            TabIndex        =   33
            Top             =   450
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   503
            _Version        =   393216
            ClipMode        =   1
            Appearance      =   0
            MaxLength       =   10
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MBox2 
            Height          =   285
            Index           =   4
            Left            =   9810
            TabIndex        =   34
            TabStop         =   0   'False
            Top             =   450
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   503
            _Version        =   393216
            ClipMode        =   1
            Appearance      =   0
            MaxLength       =   10
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MBox2 
            Height          =   285
            Index           =   5
            Left            =   10740
            TabIndex        =   35
            Top             =   450
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   503
            _Version        =   393216
            ClipMode        =   1
            Appearance      =   0
            MaxLength       =   10
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MBox2 
            Height          =   285
            Index           =   11
            Left            =   90
            TabIndex        =   113
            Top             =   450
            Width           =   585
            _ExtentX        =   1032
            _ExtentY        =   503
            _Version        =   393216
            ClipMode        =   1
            Appearance      =   0
            BackColor       =   -2147483644
            Enabled         =   0   'False
            MaxLength       =   6
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MBox2 
            Height          =   285
            Index           =   12
            Left            =   720
            TabIndex        =   40
            Top             =   450
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            Enabled         =   0   'False
            MaxLength       =   6
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MBox2 
            Height          =   255
            Index           =   13
            Left            =   90
            TabIndex        =   114
            Top             =   510
            Width           =   585
            _ExtentX        =   1032
            _ExtentY        =   450
            _Version        =   393216
            BackColor       =   -2147483648
            Enabled         =   0   'False
            MaxLength       =   6
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MBox2 
            Height          =   255
            Index           =   14
            Left            =   90
            TabIndex        =   115
            Top             =   450
            Width           =   585
            _ExtentX        =   1032
            _ExtentY        =   450
            _Version        =   393216
            ClipMode        =   1
            BackColor       =   -2147483644
            Enabled         =   0   'False
            MaxLength       =   6
            PromptChar      =   "_"
         End
         Begin VB.Label deta 
            Alignment       =   2  'Center
            Caption         =   "Item"
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
            Index           =   7
            Left            =   120
            TabIndex        =   125
            Top             =   180
            Width           =   465
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   4005
            TabIndex        =   124
            Top             =   450
            Width           =   3675
         End
         Begin VB.Label deta 
            Alignment       =   2  'Center
            Caption         =   "Cant.UM"
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
            Index           =   6
            Left            =   1590
            TabIndex        =   123
            Top             =   180
            Width           =   795
         End
         Begin VB.Label deta 
            Alignment       =   2  'Center
            Caption         =   "%Com"
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
            Index           =   5
            Left            =   10740
            TabIndex        =   122
            Top             =   180
            Width           =   975
         End
         Begin VB.Label deta 
            Alignment       =   2  'Center
            Caption         =   "Dscto"
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
            Index           =   4
            Left            =   9900
            TabIndex        =   121
            Top             =   180
            Width           =   735
         End
         Begin VB.Label deta 
            Alignment       =   2  'Center
            Caption         =   "Precio Vta"
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
            Index           =   3
            Left            =   8670
            TabIndex        =   120
            Top             =   180
            Width           =   1005
         End
         Begin VB.Label deta 
            Alignment       =   2  'Center
            Caption         =   "U.M."
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
            Index           =   2
            Left            =   7800
            TabIndex        =   119
            Top             =   180
            Width           =   675
         End
         Begin VB.Label deta 
            Alignment       =   2  'Center
            Caption         =   "Descripción"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   1
            Left            =   3870
            TabIndex        =   118
            Top             =   180
            Width           =   3885
         End
         Begin VB.Label deta 
            Alignment       =   2  'Center
            Caption         =   "Codigo"
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
            Index           =   0
            Left            =   2430
            TabIndex        =   117
            Top             =   180
            Width           =   1215
         End
         Begin VB.Label deta 
            Alignment       =   2  'Center
            Caption         =   "Cnt. Ref"
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
            Index           =   8
            Left            =   750
            TabIndex        =   116
            Top             =   180
            Width           =   765
         End
      End
      Begin VB.Frame Fr2 
         BackColor       =   &H00C9955A&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   2
         Left            =   210
         TabIndex        =   57
         Top             =   6450
         Width           =   11535
         Begin MSMask.MaskEdBox MBox2 
            Height          =   375
            Index           =   6
            Left            =   300
            TabIndex        =   58
            Top             =   60
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   0
            BackColor       =   12648447
            ForeColor       =   16777215
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MBox2 
            Height          =   375
            Index           =   7
            Left            =   2400
            TabIndex        =   59
            Top             =   60
            Width           =   1875
            _ExtentX        =   3307
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   0
            BackColor       =   12648447
            ForeColor       =   16777215
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MBox2 
            Height          =   375
            Index           =   8
            Left            =   4800
            TabIndex        =   60
            Top             =   60
            Width           =   1905
            _ExtentX        =   3360
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   0
            BackColor       =   12648447
            ForeColor       =   16777215
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MBox2 
            Height          =   375
            Index           =   9
            Left            =   7290
            TabIndex        =   61
            Top             =   60
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   0
            BackColor       =   12648447
            ForeColor       =   16777215
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MBox2 
            Height          =   375
            Index           =   10
            Left            =   9540
            TabIndex        =   62
            Top             =   60
            Width           =   1905
            _ExtentX        =   3360
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   0
            BackColor       =   12648447
            ForeColor       =   16777215
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00808080&
            Index           =   7
            X1              =   9340
            X2              =   9340
            Y1              =   120
            Y2              =   1215
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00808080&
            Index           =   6
            X1              =   6940
            X2              =   6940
            Y1              =   120
            Y2              =   1215
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00808080&
            Index           =   5
            X1              =   4420
            X2              =   4420
            Y1              =   120
            Y2              =   1215
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00808080&
            Index           =   4
            X1              =   2160
            X2              =   2160
            Y1              =   120
            Y2              =   1215
         End
         Begin VB.Line Line2 
            BorderColor     =   &H80000005&
            Index           =   3
            X1              =   9360
            X2              =   9360
            Y1              =   120
            Y2              =   1215
         End
         Begin VB.Line Line2 
            BorderColor     =   &H80000005&
            Index           =   2
            X1              =   6960
            X2              =   6960
            Y1              =   120
            Y2              =   1215
         End
         Begin VB.Line Line2 
            BorderColor     =   &H80000005&
            Index           =   1
            X1              =   4440
            X2              =   4440
            Y1              =   120
            Y2              =   1215
         End
         Begin VB.Line Line2 
            BorderColor     =   &H80000005&
            Index           =   0
            X1              =   2175
            X2              =   2175
            Y1              =   120
            Y2              =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Neto Factura"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   4
            Left            =   9840
            TabIndex        =   67
            Top             =   495
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Total I.G.V."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   3
            Left            =   7680
            TabIndex        =   66
            Top             =   495
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Total Dctos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   2
            Left            =   5160
            TabIndex        =   65
            Top             =   495
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Total Bruto"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   1
            Left            =   2760
            TabIndex        =   64
            Top             =   495
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Cantidad"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   63
            Top             =   495
            Width           =   1335
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
         Left            =   240
         TabIndex        =   52
         Top             =   7170
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
            Index           =   0
            Left            =   1470
            TabIndex        =   56
            Top             =   270
            Width           =   1095
         End
         Begin VB.Label Label4 
            Caption         =   "Editar Item"
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
            Index           =   1
            Left            =   1470
            TabIndex        =   55
            Top             =   570
            Width           =   1125
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   150
            Picture         =   "FrmCotizacionLibre.frx":0948
            Top             =   330
            Width           =   480
         End
         Begin VB.Label Label5 
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
            Index           =   0
            Left            =   780
            TabIndex        =   54
            Top             =   270
            Width           =   495
         End
         Begin VB.Label Label5 
            Caption         =   "[ENTER]"
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
            Index           =   1
            Left            =   780
            TabIndex        =   53
            Top             =   570
            Width           =   675
         End
      End
      Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
         Height          =   2595
         Left            =   60
         TabIndex        =   68
         Top             =   3810
         Width           =   11835
         _ExtentX        =   20876
         _ExtentY        =   4577
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
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
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
      Begin Crystal.CrystalReport oCrystalReport 
         Left            =   630
         Top             =   7200
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   2535
         Left            =   30
         TabIndex        =   69
         Top             =   390
         Width           =   11805
         _ExtentX        =   20823
         _ExtentY        =   4471
         _Version        =   393216
         Style           =   1
         Tab             =   1
         TabHeight       =   520
         TabCaption(0)   =   "Datos Generales"
         TabPicture(0)   =   "FrmCotizacionLibre.frx":0D8A
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Fr1(0)"
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Datos Detalle"
         TabPicture(1)   =   "FrmCotizacionLibre.frx":0DA6
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "TClie"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "Fr2(0)"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).ControlCount=   2
         TabCaption(2)   =   "Datos Complementarios"
         TabPicture(2)   =   "FrmCotizacionLibre.frx":0DC2
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Fr3(0)"
         Tab(2).ControlCount=   1
         Begin VB.Frame Fr3 
            BackColor       =   &H80000000&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1755
            Index           =   0
            Left            =   -74880
            TabIndex        =   99
            Top             =   450
            Width           =   11565
            Begin VB.ComboBox Combo6 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   7320
               Style           =   2  'Dropdown List
               TabIndex        =   47
               Top             =   930
               Width           =   1125
            End
            Begin VB.ComboBox Combo7 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   9540
               Style           =   2  'Dropdown List
               TabIndex        =   48
               Top             =   930
               Width           =   1410
            End
            Begin VB.ComboBox Combo8 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   1290
               Style           =   2  'Dropdown List
               TabIndex        =   49
               Top             =   1290
               Width           =   1185
            End
            Begin MSMask.MaskEdBox MBox3 
               Height          =   315
               Index           =   0
               Left            =   1290
               TabIndex        =   42
               Top             =   210
               Width           =   1395
               _ExtentX        =   2461
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   11
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MBox3 
               Height          =   315
               Index           =   1
               Left            =   2745
               TabIndex        =   43
               Top             =   210
               Width           =   6495
               _ExtentX        =   11456
               _ExtentY        =   556
               _Version        =   393216
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MBox3 
               Height          =   315
               Index           =   2
               Left            =   9840
               TabIndex        =   44
               Top             =   210
               Width           =   1635
               _ExtentX        =   2884
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   11
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MBox3 
               Height          =   315
               Index           =   3
               Left            =   1290
               TabIndex        =   45
               Top             =   570
               Width           =   10185
               _ExtentX        =   17965
               _ExtentY        =   556
               _Version        =   393216
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MBox3 
               Height          =   315
               Index           =   4
               Left            =   1290
               TabIndex        =   46
               Top             =   930
               Width           =   4545
               _ExtentX        =   8017
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   20
               PromptChar      =   "_"
            End
            Begin VB.Label Dclie 
               BackStyle       =   0  'Transparent
               Caption         =   "Limite Cred US$"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000006&
               Height          =   255
               Index           =   5
               Left            =   8520
               TabIndex        =   110
               Top             =   1380
               Width           =   1815
            End
            Begin VB.Label Dclie 
               BackStyle       =   0  'Transparent
               Caption         =   "Saldo US$"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000006&
               Height          =   255
               Index           =   4
               Left            =   5790
               TabIndex        =   109
               Top             =   1380
               Width           =   1335
            End
            Begin VB.Label Dclie 
               BackStyle       =   0  'Transparent
               Caption         =   "Distrito"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000006&
               Height          =   255
               Index           =   2
               Left            =   240
               TabIndex        =   108
               Top             =   990
               Width           =   1815
            End
            Begin VB.Label Dclie 
               BackStyle       =   0  'Transparent
               Caption         =   "Direccion"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000006&
               Height          =   255
               Index           =   1
               Left            =   240
               TabIndex        =   107
               Top             =   600
               Width           =   1815
            End
            Begin VB.Label Dclie 
               BackStyle       =   0  'Transparent
               Caption         =   "Cliente"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000006&
               Height          =   255
               Index           =   0
               Left            =   240
               TabIndex        =   106
               Top             =   240
               Width           =   1815
            End
            Begin VB.Label lcred 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000C&
               Height          =   285
               Index           =   0
               Left            =   6780
               TabIndex        =   105
               Top             =   1350
               Width           =   1575
            End
            Begin VB.Label lcred 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000E&
               Height          =   285
               Index           =   1
               Left            =   9870
               TabIndex        =   104
               Top             =   1320
               Width           =   1605
            End
            Begin VB.Label Dclie 
               BackStyle       =   0  'Transparent
               Caption         =   "Ruc"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000006&
               Height          =   255
               Index           =   7
               Left            =   9420
               TabIndex        =   103
               Top             =   270
               Width           =   675
            End
            Begin VB.Label Dclie 
               BackStyle       =   0  'Transparent
               Caption         =   "Tipo Persona"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000006&
               Height          =   255
               Index           =   3
               Left            =   6120
               TabIndex        =   102
               Top             =   960
               Width           =   975
            End
            Begin VB.Label Dclie 
               BackStyle       =   0  'Transparent
               Caption         =   "Pais"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000006&
               Height          =   255
               Index           =   6
               Left            =   9030
               TabIndex        =   101
               Top             =   990
               Width           =   465
            End
            Begin VB.Label Dclie 
               BackStyle       =   0  'Transparent
               Caption         =   "Multidireccion"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000006&
               Height          =   255
               Index           =   8
               Left            =   240
               TabIndex        =   100
               Top             =   1380
               Width           =   1005
            End
         End
         Begin VB.Frame Fr2 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1875
            Index           =   0
            Left            =   60
            TabIndex        =   83
            Top             =   330
            Width           =   11685
            Begin VB.ComboBox Combo5 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   9300
               Style           =   2  'Dropdown List
               TabIndex        =   28
               Top             =   1170
               Width           =   735
            End
            Begin VB.ComboBox Combo4 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   9120
               Style           =   2  'Dropdown List
               TabIndex        =   20
               Top             =   180
               Width           =   2445
            End
            Begin VB.ComboBox Combo3 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   1890
               Style           =   2  'Dropdown List
               TabIndex        =   18
               Top             =   180
               Width           =   2265
            End
            Begin VB.CommandButton cAyuda 
               Caption         =   "..."
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   0
               Left            =   9045
               TabIndex        =   37
               Top             =   1530
               Width           =   285
            End
            Begin VB.TextBox Text1 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   10530
               TabIndex        =   38
               Top             =   1530
               Visible         =   0   'False
               Width           =   1005
            End
            Begin VB.TextBox Text2 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   60
               TabIndex        =   84
               Top             =   1800
               Visible         =   0   'False
               Width           =   1005
            End
            Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuda3 
               Height          =   315
               Left            =   8940
               TabIndex        =   24
               Top             =   840
               Width           =   2625
               _ExtentX        =   4630
               _ExtentY        =   556
               XcodMaxLongitud =   2
               xcodwith        =   100
               NomTabla        =   "tabalm"
               TituloAyuda     =   "Ayuda de Almacenes"
               ListaCampos     =   "taalma(1),tadescri(1)"
               XcodCampo       =   "taalma"
               XListCampo      =   "tadescri"
               ListaCamposDescrip=   "Codigo,Descripcion"
               ListaCamposText =   "taalma,tadescri"
            End
            Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuda2 
               Height          =   315
               Left            =   1890
               TabIndex        =   22
               Top             =   825
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
            End
            Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuda1 
               Height          =   315
               Left            =   1890
               TabIndex        =   21
               Top             =   510
               Width           =   9645
               _ExtentX        =   17013
               _ExtentY        =   556
               XcodMaxLongitud =   11
               xcodwith        =   800
               NomTabla        =   "vt_Cliente"
               TituloAyuda     =   "Ayuda de Clientes"
               ListaCampos     =   $"FrmCotizacionLibre.frx":0DDE
               XcodCampo       =   "clientecodigo"
               XListCampo      =   "clienterazonsocial"
               ListaCamposDescrip=   "Codigo,Descripcion,Ruc,Direccion,Distrito,LimiteCred,Saldo,T,P,M,D"
               ListaCamposText =   $"FrmCotizacionLibre.frx":0EC4
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
               Left            =   6060
               TabIndex        =   19
               Top             =   180
               Width           =   1155
               _ExtentX        =   2037
               _ExtentY        =   503
               _Version        =   393216
               ClipMode        =   1
               Appearance      =   0
               AllowPrompt     =   -1  'True
               MaxLength       =   10
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MBox 
               Height          =   255
               Index           =   13
               Left            =   6750
               TabIndex        =   23
               Top             =   840
               Width           =   1185
               _ExtentX        =   2090
               _ExtentY        =   450
               _Version        =   393216
               ClipMode        =   1
               Appearance      =   0
               MaxLength       =   10
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MBox 
               Height          =   255
               Index           =   15
               Left            =   1890
               TabIndex        =   25
               Top             =   1200
               Width           =   1185
               _ExtentX        =   2090
               _ExtentY        =   450
               _Version        =   393216
               ClipMode        =   1
               Appearance      =   0
               MaxLength       =   10
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MBox 
               Height          =   255
               Index           =   16
               Left            =   4470
               TabIndex        =   26
               Top             =   1200
               Width           =   1185
               _ExtentX        =   2090
               _ExtentY        =   450
               _Version        =   393216
               ClipMode        =   1
               Appearance      =   0
               MaxLength       =   10
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MBox 
               Height          =   255
               Index           =   17
               Left            =   7110
               TabIndex        =   27
               Top             =   1200
               Width           =   1035
               _ExtentX        =   1826
               _ExtentY        =   450
               _Version        =   393216
               ClipMode        =   1
               Appearance      =   0
               MaxLength       =   10
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MBox 
               Height          =   255
               Index           =   18
               Left            =   11010
               TabIndex        =   29
               Top             =   1200
               Width           =   525
               _ExtentX        =   926
               _ExtentY        =   450
               _Version        =   393216
               ClipMode        =   1
               Appearance      =   0
               MaxLength       =   3
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MBox 
               Height          =   285
               Index           =   19
               Left            =   1890
               TabIndex        =   30
               Top             =   1500
               Width           =   7110
               _ExtentX        =   12541
               _ExtentY        =   503
               _Version        =   393216
               Appearance      =   0
               PromptChar      =   "_"
            End
            Begin VB.Label Label1 
               Caption         =   "% Comision"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   23
               Left            =   5850
               TabIndex        =   98
               Top             =   870
               Width           =   855
            End
            Begin VB.Label Label1 
               Caption         =   "Autorizacion"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   22
               Left            =   8310
               TabIndex        =   97
               Top             =   1200
               Width           =   945
            End
            Begin VB.Label Label1 
               Caption         =   "Orden de Compra"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   21
               Left            =   5760
               TabIndex        =   96
               Top             =   1170
               Width           =   1395
            End
            Begin VB.Label Label1 
               Caption         =   "Nota de Pedido"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   20
               Left            =   3240
               TabIndex        =   95
               Top             =   1170
               Width           =   1575
            End
            Begin VB.Label Label1 
               Caption         =   "Otros Gastos"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   19
               Left            =   150
               TabIndex        =   94
               Top             =   1200
               Width           =   1575
            End
            Begin VB.Label Label1 
               Caption         =   "Almacen"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   17
               Left            =   8130
               TabIndex        =   93
               Top             =   900
               Width           =   795
            End
            Begin VB.Label Label1 
               Caption         =   "Codigo del Vendedor"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   16
               Left            =   150
               TabIndex        =   92
               Top             =   900
               Width           =   1575
            End
            Begin VB.Label Label1 
               Caption         =   "Codigo del Cliente"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   15
               Left            =   150
               TabIndex        =   91
               Top             =   570
               Width           =   1575
            End
            Begin VB.Label Label1 
               Caption         =   "Forma de Pago"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   14
               Left            =   7920
               TabIndex        =   90
               Top             =   240
               Width           =   1095
            End
            Begin VB.Label Label1 
               Caption         =   "Fecha de Atencion"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   13
               Left            =   4590
               TabIndex        =   89
               Top             =   240
               Width           =   1365
            End
            Begin VB.Label Label1 
               Caption         =   "Modo de la Venta"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   12
               Left            =   150
               TabIndex        =   88
               Top             =   240
               Width           =   1575
            End
            Begin VB.Label Label1 
               Caption         =   "Dias Pago"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   18
               Left            =   10170
               TabIndex        =   87
               Top             =   1200
               Width           =   765
            End
            Begin VB.Label Label1 
               Caption         =   "Punto de Llegada"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   24
               Left            =   120
               TabIndex        =   86
               Top             =   1530
               Width           =   1335
            End
            Begin VB.Label Label6 
               Caption         =   "Dscto Cliente"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   9525
               TabIndex        =   85
               Top             =   1545
               Visible         =   0   'False
               Width           =   975
            End
         End
         Begin VB.Frame Fr1 
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1695
            Index           =   0
            Left            =   -74910
            TabIndex        =   70
            Top             =   570
            Width           =   11565
            Begin VB.TextBox TxtContacto 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   3480
               MaxLength       =   200
               TabIndex        =   17
               Top             =   1020
               Width           =   8055
            End
            Begin VB.ComboBox Combo2 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   1290
               Style           =   2  'Dropdown List
               TabIndex        =   16
               Top             =   1020
               Width           =   1065
            End
            Begin VB.ComboBox Combo1 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   7920
               Style           =   2  'Dropdown List
               TabIndex        =   14
               Top             =   630
               Width           =   1305
            End
            Begin MSMask.MaskEdBox MBox 
               Height          =   285
               Index           =   0
               Left            =   1290
               TabIndex        =   6
               Top             =   240
               Width           =   435
               _ExtentX        =   767
               _ExtentY        =   503
               _Version        =   393216
               Appearance      =   0
               Enabled         =   0   'False
               MaxLength       =   2
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MBox 
               Height          =   285
               Index           =   1
               Left            =   3030
               TabIndex        =   7
               Top             =   240
               Width           =   1245
               _ExtentX        =   2196
               _ExtentY        =   503
               _Version        =   393216
               Appearance      =   0
               MaxLength       =   11
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MBox 
               Height          =   285
               Index           =   2
               Left            =   5460
               TabIndex        =   8
               Top             =   240
               Width           =   1125
               _ExtentX        =   1984
               _ExtentY        =   503
               _Version        =   393216
               ClipMode        =   1
               Appearance      =   0
               Enabled         =   0   'False
               MaxLength       =   11
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MBox 
               Height          =   285
               Index           =   3
               Left            =   7770
               TabIndex        =   9
               Top             =   240
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   503
               _Version        =   393216
               ClipMode        =   1
               Appearance      =   0
               Enabled         =   0   'False
               MaxLength       =   11
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MBox 
               Height          =   285
               Index           =   4
               Left            =   10140
               TabIndex        =   10
               Top             =   240
               Width           =   1395
               _ExtentX        =   2461
               _ExtentY        =   503
               _Version        =   393216
               ClipMode        =   1
               Appearance      =   0
               Enabled         =   0   'False
               MaxLength       =   11
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MBox 
               Height          =   285
               Index           =   5
               Left            =   1290
               TabIndex        =   11
               Top             =   630
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   503
               _Version        =   393216
               ClipMode        =   1
               Appearance      =   0
               MaxLength       =   9
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MBox 
               Height          =   285
               Index           =   6
               Left            =   3570
               TabIndex        =   12
               Top             =   630
               Width           =   1035
               _ExtentX        =   1826
               _ExtentY        =   503
               _Version        =   393216
               ClipMode        =   1
               Appearance      =   0
               MaxLength       =   9
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MBox 
               Height          =   285
               Index           =   7
               Left            =   5970
               TabIndex        =   13
               Top             =   630
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   503
               _Version        =   393216
               ClipMode        =   1
               Appearance      =   0
               MaxLength       =   9
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MBox 
               Height          =   285
               Index           =   8
               Left            =   10560
               TabIndex        =   15
               Top             =   630
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   503
               _Version        =   393216
               ClipMode        =   1
               Appearance      =   0
               MaxLength       =   9
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MBox 
               Height          =   135
               Index           =   9
               Left            =   120
               TabIndex        =   36
               Top             =   1770
               Visible         =   0   'False
               Width           =   765
               _ExtentX        =   1349
               _ExtentY        =   238
               _Version        =   393216
               MaxLength       =   45
               PromptChar      =   "_"
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Contacto :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   11
               Left            =   2610
               TabIndex        =   82
               Top             =   1050
               Width           =   735
            End
            Begin VB.Label Label1 
               Caption         =   "Lista Precios"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   10
               Left            =   120
               TabIndex        =   81
               Top             =   1050
               Width           =   1215
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Dcto. Especial"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   9
               Left            =   4800
               TabIndex        =   80
               Top             =   660
               Width           =   1035
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Nro. Guia :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   8
               Left            =   9270
               TabIndex        =   79
               Top             =   270
               Width           =   765
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Nro. Pedido :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   7
               Left            =   2010
               TabIndex        =   78
               Top             =   270
               Width           =   930
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Moneda"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   6
               Left            =   7140
               TabIndex        =   77
               Top             =   660
               Width           =   585
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Dcto. Promoc."
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   5
               Left            =   2490
               TabIndex        =   76
               Top             =   660
               Width           =   1020
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Nro .Boleta :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   4
               Left            =   6780
               TabIndex        =   75
               Top             =   270
               Width           =   885
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Tipo de Cambio"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   3
               Left            =   9330
               TabIndex        =   74
               Top             =   660
               Width           =   1110
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Dcto. Genral."
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   2
               Left            =   120
               TabIndex        =   73
               Top             =   660
               Width           =   945
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Nro .Factura :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   1
               Left            =   4380
               TabIndex        =   72
               Top             =   270
               Width           =   975
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Punto Venta :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   71
               Top             =   270
               Width           =   975
            End
         End
         Begin VB.CheckBox TClie 
            Caption         =   "Cliente Eventual"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   9960
            TabIndex        =   39
            Top             =   2250
            Width           =   1515
         End
      End
      Begin TrueOleDBGrid70.TDBGrid TDBGrid2 
         Height          =   5850
         Left            =   -74820
         TabIndex        =   128
         Top             =   1710
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   10319
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
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
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
      Begin MSComCtl2.DTPicker DtFechaHasta 
         Height          =   285
         Left            =   -71355
         TabIndex        =   3
         Top             =   1305
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   503
         _Version        =   393216
         Format          =   117047297
         CurrentDate     =   39763
         MaxDate         =   44196
         MinDate         =   36526
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Razon Social :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74775
         TabIndex        =   135
         Top             =   990
         Width           =   1155
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Fecha :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74775
         TabIndex        =   134
         Top             =   1350
         Width           =   585
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Nro Cotizacion :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74775
         TabIndex        =   133
         Top             =   630
         Width           =   1290
      End
      Begin VB.Label LblReg 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "(0) Cotizaciones"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -64740
         TabIndex        =   132
         Top             =   7695
         Width           =   1365
      End
   End
End
Attribute VB_Name = "FrmCotizacionLibre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
                    
Option Explicit

Dim nLongicampo(6) As Integer
Dim rsdeta As New ADODB.Recordset
Dim wCabe(40)

'****** Totales de Pedidos***
Dim Tbruto As Double
Dim Tigv As Double
Dim Tdscto As Double
Dim TSub As Double
Dim TImporte As Double
Dim TNeto As Double
Dim TCant As Double
Dim flag As Integer

'***Total Descuentos  ***
Dim DTGlobal As Double
Dim DTCliente As Double
Dim DTPPago As Double
Dim DTOficina As Double
Dim DTItem As Double
Dim DTLinea As Double
Dim DTPromo As Double

'*****************

Dim dllgeneral As New dllgeneral.dll_general

'Mensajes de Pedidos

Const W1TXT1 = "El Cliente No Existe en el Maestro de Clientes"
Const W1TXT2 = "El Cliente No Tiene Número de R.U.C. en el Maestro"
Const W1TXT3 = "El Cliente Esta Suspendido No Atender"
Const W1TXT4 = "El Cliente Ya No Tiene Credito. No Atender"

Const W1TXT6 = "Codigo del Vendedor No Existe en Tabla de Vendedores"
Const W1TXT7 = "El Codigo del Almacen No Existe en Tabla de Almacenes"

Const W1TXT9 = "El Monto de Otros Gastos debe ser un Valor Positivo"

Const W1TXT12 = "El Descuento General debe ser un Valor Positivo"
Const W1TXT13 = "El Descuento de Promoci¢n debe ser un Valor Positivo"
Const W1TXT14 = "El Descuento Pronto Pago debe ser un Valor Positivo"
Const W1TXT17 = "Codigo de la Lista de Precios No Existe"
Const W1TXT18 = "Archivo Maestro de la Lista de Precios No Existe"
Const W1TXT19 = "Codigo del Artículo No Existe en Maestro de Artículos "
Const W1TXT20 = "El Codigo del Articulo No Existe en Maestro de Precios"
Const W1TXT21 = "El Codigo del Articulo Ya Existe en el Proceso de Ventas"
Const W1TXT22 = "La Cantidad a Vender debe ser un Valor Mayor que Cero"
Const W1TXT23 = "La Cantidad a Vender es Mayor que el Actual en Almacén"
Const W1TXT24 = "El Precio de Venta debe de ser un Valor Mayor que Cero"
Const W1TXT25 = "El Descuento por Item debe ser un Valor Positivo"
Const W1TXT28 = "Debe de Ingresar el Nro. de R.U.C. del Cliente"
Const W1TXT30 = "El Importe debe ser mayor a cero"
Const W1TXT31 = ""
Const W1TXT32 = ""
Const W1TXT33 = ""


Private Sub cAyuda_Click(Index As Integer)
  If Index = 0 And Len(Trim(MBox(19))) = 0 Then    'Ayuda de Punto de LLegada
    If dllgeneral.VerificaDatoExistente(VGCNx, "select * from vt_cliente where clientecodigo='" & Ctr_Ayuda1.xclave & "'") = 1 Then
       Dim gfiltra(1, 2) As String
       gfiltra(1, 1) = "Descripcion": gfiltra(1, 2) = "clientedireccion"
       FrmAyudaPedidos.TipoForma = 1
       FrmAyudaPedidos.BConexion = VGCNx
       FrmAyudaPedidos.Bdata = "0"
       FrmAyudaPedidos.BTabla = "vt_clientedireccion"
       FrmAyudaPedidos.BCampos = "Clientecodigo as Codigo,Cliedirdireccion as Descripcion"
       FrmAyudaPedidos.BOrden = "Cliedirdireccion"
       FrmAyudaPedidos.BCondi = "clientecodigo='" & Ctr_Ayuda1.xclave & "'"
       FrmAyudaPedidos.BFiltro = gfiltra
    Else
        nAyuda = "": nDetalle = ""
        MsgBox "No existen Direcciones Anexas...", vbInformation, MsgTitle
        Exit Sub
    End If
  ElseIf Index = 3 Then                             ' Ayuda de Productos
'       If Len(Label2) > 0 Then
'         SendKeys "{tab}"
'         Exit Sub
'       End If
       
       Dim sfiltra(1 To 2, 1 To 2) As String
       sfiltra(1, 1) = "Codigo": sfiltra(1, 2) = "acodigo"
       sfiltra(2, 1) = "Descripcion": sfiltra(2, 2) = "adescri"
       FrmAyudaPedidos.TipoForma = 1
       FrmAyudaPedidos.BConexion = VGCNx
       If Combo2.ListCount > 0 Then
          FrmAyudaPedidos.BTabla = "[" & VGCNx.DefaultDatabase & "].dbo.maeart inner join [" & _
                            VGCNx.DefaultDatabase & "].dbo.stkart " & _
                            " ON acodigo=stcodigo"
       Else
          FrmAyudaPedidos.BTabla = "[" & VGCNx.DefaultDatabase & "].dbo.maeart inner join [" & _
                            VGCNx.DefaultDatabase & "].dbo.stkart " & _
                            " ON acodigo=stcodigo "

       End If
       FrmAyudaPedidos.Bdata = "3"
       FrmAyudaPedidos.Bdato = Escadena(MBox2(1).Text)
       FrmAyudaPedidos.BCampos = "acodigo as Codigo,adescri as Descripcion,stskdis as Stock"
       FrmAyudaPedidos.BOrden = "adescri"
       FrmAyudaPedidos.BCondi = "stalma='" & Ctr_Ayuda3.xclave & "' and stskdis>0"
       FrmAyudaPedidos.BFiltro = sfiltra
   Else
       SendKeys "{tab}"
       Exit Sub
   End If
   FrmAyudaPedidos.Show 1
   If Index = 3 Then
       MBox2(1) = Escadena(nAyuda):   Label2 = Escadena(nDetalle)
   ElseIf Index = 0 Then
       MBox(19) = Escadena(nDetalle)
   End If
   nAyuda = "": nDetalle = ""
End Sub

Private Sub cBoton_Click(Index As Integer)
  Dim J As Integer
  Dim valor As Double
  If Index = 0 Then
'       Fr1(1).Visible = False
        TClie.Value = 0
       Limpiartexto MBox, 2, 9
       MBox(0).Enabled = False:  MBox(1).Enabled = False
       MBox(0).Text = g_ptoventa
       valor = 0
       If TDBGrid2.ApproxCount > 0 Then
            TDBGrid2.MoveFirst
            Do Until TDBGrid2.EOF
                valor = CDbl(TDBGrid2.Columns(0).Text)
                TDBGrid2.MoveNext
            Loop
       Else
            valor = 1
       End If
       MBox(1) = Right("000000000000" & Trim(CStr(valor)), MBox(1).MaxLength)
       MBox(2) = Right("000000000000", MBox(2).MaxLength)
       MBox(3) = Right("000000000000", MBox(3).MaxLength)
       MBox(4) = Right("000000000000", MBox(4).MaxLength)
       MBox(5) = numero(0): MBox(6) = numero(0): MBox(7) = numero(0): MBox(8) = numero(TraeTipoCambio(Date, VGCNx))
       TxtContacto.Text = Escadena(VGParamSistem.mensaje)
       MBox(19) = ""
       MBox(10) = Format(Date, "dd/mm/yyyy")
       MBox(13) = numero(0)
       MBox(15) = numero(0)
       MBox(16) = 0: MBox(17) = 0: MBox(18) = 0
       For J = 0 To 5
          MBox2(J) = ""
       Next J
       Set rsdeta = Nothing
       
       CargaGrilla

     'Se activa los parametros deventa
       Combo1.ListIndex = VerificaCombo(Combo1, VGParamSistem.moneda)     'moneda
       Combo2.ListIndex = VerificaCombo(Combo2, VGParamSistem.listapre)   'listaprecios
       Combo2.Enabled = False
       MBox(8) = numero(VGParamSistem.tipocambio)                         'tipo de cambio
       Ctr_Ayuda3.xclave = Escadena(VGParamSistem.almacen)                'almacen
       Call Ctr_Ayuda3.Ejecutar
       MBox(13).Enabled = IIf(VGParamSistem.comivende = "F", False, True)                     'comision de vendedor
       
      'Se activa los parametros de punto de venta
       MBox(2).Enabled = IIf(VGParametros.nrofactura = "1" And VGParametros.ventaauto = "0", True, False)
       MBox(3).Enabled = IIf(VGParametros.nroboleta = "1" And VGParametros.ventaauto = "0", True, False)
       MBox(4).Enabled = IIf(VGParametros.nroguia = "1" And VGParametros.ventaauto = "0", True, False)
       
     'Activamos el Tab
       Activa 1
       SSTab2.TabEnabled(2) = False
       SSTab2.Tab = 0
       MBox(5).SetFocus

  ElseIf Index = 1 Then
      Fr1(1).Visible = False
  End If
End Sub





Private Sub cmdBotones_Click(Index As Integer)
   Dim asql As String
   Dim acmd As New ADODB.Command
   Dim J As Integer
   
   On Error Resume Next 'GoTo vererror
   
   Select Case Index
    Case 0
'        Fr1(1).Visible = True
'        Limpiartexto MBox2, 6, 10
'        Fr1(0).Enabled = True
'        Fr2(0).Enabled = True
'        Fr3(0).Enabled = True
        TClie.Enabled = True
        g_TipoMovi = 1
        MBox(0).Text = g_ptoventa
       MBox(1) = g_pedserie & Right("000000000000" & TraeDataSerie("select puntovtadoccorr from vt_puntovtadocumento where documentocodigo='" & g_tipoped & "' and puntovtadocserie='" & g_pedserie & "' and puntovtacodigo='" & g_ptoventa & "' and empresacodigo='" & VGParametros.empresacodigo & "'", VGCNx), 8)
       MBox(2) = g_facserie & Right("000000000000" & TraeDataSerie("select puntovtadoccorr from vt_puntovtadocumento where documentocodigo='" & g_tipofac & "' and puntovtadocserie='" & g_facserie & "' and puntovtacodigo='" & g_ptoventa & "' and empresacodigo='" & VGParametros.empresacodigo & "'", VGCNx), 8)
       MBox(3) = g_bolserie & Right("000000000000" & TraeDataSerie("select puntovtadoccorr from vt_puntovtadocumento where documentocodigo='" & g_tipobol & "' and puntovtadocserie='" & g_bolserie & "' and puntovtacodigo='" & g_ptoventa & "' and empresacodigo='" & VGParametros.empresacodigo & "'", VGCNx), 8)
       MBox(4) = g_guiaserie & Right("000000000000" & TraeDataSerie("select puntovtadoccorr from vt_puntovtadocumento where documentocodigo='" & g_tipoguia & "' and puntovtadocserie='" & g_guiaserie & "' and puntovtacodigo='" & g_ptoventa & "' and empresacodigo='" & VGParametros.empresacodigo & "'", VGCNx), 8)
       MBox(5) = numero(0): MBox(6) = numero(0): MBox(7) = numero(0): MBox(8) = numero(TraeTipoCambio(Date, VGCNx))
 '       Call cBoton_Click(0)
        Activa 1
        
    Case 1
       If TDBGrid2.Row >= 0 Then
          Fr1(0).Enabled = True
          Fr2(0).Enabled = True
          Fr3(0).Enabled = True
          TClie.Enabled = True
          Limpiartexto MBox2, 6, 10
          Call Carga_Pedido
          Activa 1
       End If
    Case 2
       If TDBGrid2.Row >= 0 Then
        asql = "pedidonumero='" & TDBGrid2.Columns(0).Text & "'"
        If dllgeneral.EliminaReg(VGCNx, "detallecotizalibre", asql) = 1 Then
            VGCNx.Execute "Delete From cotizalibre where " & asql
        End If
        Listado
       End If
    Case 4
       Unload Me
    Case 11
        If IsNull(Ctr_Ayuda1.xclave) Or Len(Trim(Ctr_Ayuda1.xclave)) = 0 Then
           MsgBox W1TXT1, vbInformation, MsgTitle
           SSTab2.Tab = 1
           Ctr_Ayuda1.SetFocus
           Exit Sub
        End If
        If IsNull(Ctr_Ayuda2.xclave) Or Len(Trim(Ctr_Ayuda2.xclave)) = 0 Then
           MsgBox W1TXT6, vbInformation, MsgTitle
           SSTab2.Tab = 1
           Ctr_Ayuda2.SetFocus
           Exit Sub
        End If
        If IsNull(Ctr_Ayuda3.xclave) Or Len(Trim(Ctr_Ayuda3.xclave)) = 0 Then
           MsgBox W1TXT7, vbInformation, MsgTitle
           SSTab2.Tab = 1
           Ctr_Ayuda3.SetFocus
           Exit Sub
        End If
        If IsNull(MBox(8)) Or Len(Trim(MBox(8))) = 0 Or CDbl(MBox(8)) <= 0 Then
           MsgBox "Falta Tipo de Cambio", vbInformation, MsgTitle
           Call dllgeneral.Enfoquetexto(MBox(8))
           Exit Sub
        End If
        If IsNull(MBox(15)) Or Len(Trim(MBox(15))) = 0 Or CDbl(MBox(15)) < 0 Then
           MsgBox W1TXT9, vbInformation, MsgTitle
           SSTab2.Tab = 1
           Exit Sub
        End If
        
        VGCNx.BeginTrans
        If GrabarData() = 1 Then
           VGCNx.CommitTrans
           Call CotizaImprimir
        Else
           VGCNx.RollbackTrans
        End If
        g_TipoMovi = 0
        Listado
        Activa 2
        Exit Sub

    Case 12
       Activa 2
       g_TipoMovi = 0
   End Select
   
vererror:
    If Err Then
       MsgBox "Comunicarse con Sistemas ...!!" & Chr(13) & Chr(10) & VGCNx.Errors(0).Number & "-" & VGCNx.Errors(0).Description
       Err = 0
       'VGcnx.RollbackTrans
       Exit Sub
    End If
End Sub

Public Function Activa(ntipo As Integer)
    If ntipo = 1 Then
        SSTab1.TabEnabled(0) = False
        SSTab1.TabEnabled(1) = True
        SSTab1.Tab = 1
    ElseIf ntipo = 2 Then
        SSTab1.TabEnabled(0) = True
        SSTab1.TabEnabled(1) = False
        SSTab1.Tab = 0
    End If
End Function

Private Sub CmdBuscar_Click()
Listado
TxtNro.SetFocus
End Sub

Private Sub Combo1_Click()
'   MBox(8) = Numero(0) ' Numero(TraeDataSerie("select * from ct_tipocambio where tipocambiofecha=GETDATE()",VGcnxconta))
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
  Seguir Combo1, KeyAscii
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
  Seguir Combo2, KeyAscii
End Sub

Private Sub Combo3_Click()
  Dim rs As New ADODB.Recordset
  If Combo3.ListCount > 0 Then
     Set rs = VGCNx.Execute("select * from vt_modoventa where modovtacodigo='" & dllgeneral.ComboDato(Combo3.Text) & "'")
     If rs.RecordCount > 0 Then
        modoventa.descuento = Escadena(rs!modovtadscto)
        modoventa.impuestos = Escadena(IIf(IsNull(rs!modovtaimpuestos) Or rs!modovtaimpuestos = 0, "0", "1"))
        modoventa.nroitem = IIf(IsNull(rs!modovtaitemxdoc), 10, rs!modovtaitemxdoc)
        modoventa.copiashoja = IIf(IsNull(rs!modovtacopiashojatrab), 1, rs!modovtacopiashojatrab)
        modoventa.copiasbol = IIf(IsNull(rs!modovtacopiasboleta), 1, rs!modovtacopiasboleta)
        modoventa.copiasfac = IIf(IsNull(rs!modovtacopiasfact), 1, rs!modovtacopiasfact)
        modoventa.ctacte = Escadena(IIf(IsNull(rs!modovtaactctacte) Or rs!modovtaactctacte = 0, "0", "1"))
        modoventa.ctrlinventario = Escadena(IIf(IsNull(rs!modovtactrlinventario) Or rs!modovtactrlinventario = 0, "0", "1"))
        modoventa.emitehoja = Escadena(IIf(IsNull(rs!modovtaemitehoja) Or rs!modovtaemitehoja = 0, "0", "1"))
        modoventa.emitefact = Escadena(IIf(IsNull(rs!modovtasolemitfact) Or rs!modovtasolemitfact = 0, "0", "1"))
        modoventa.emiteguia = Escadena(IIf(IsNull(rs!modovtaemiteguia) Or rs!modovtaemiteguia = 0, "0", "1"))
        modoventa.ingcliente = Escadena(IIf(IsNull(rs!modovtaingcodclie) Or rs!modovtaingcodclie = 0, "0", "1"))
        modoventa.ingforma = Escadena(IIf(IsNull(rs!modovtaingformapag) Or rs!modovtaingformapag = 0, "0", "1"))
        modoventa.ingguia = Escadena(IIf(IsNull(rs!modovtaingguiarem) Or rs!modovtaingguiarem = 0, "0", "1"))
        modoventa.inghoja = Escadena(IIf(IsNull(rs!modovtainghojatrab) Or rs!modovtainghojatrab = 0, "0", "1"))
        modoventa.ingpedido = Escadena(IIf(IsNull(rs!modovtaingpedido) Or rs!modovtaingpedido = 0, "0", "1"))
        modoventa.modificaguia = Escadena(IIf(IsNull(rs!modovtacorrguiarem) Or rs!modovtacorrguiarem = 0, "0", "1"))
        modoventa.unidadmedida = Escadena(IIf(IsNull(rs!modovtaunidadmedida), "V", Escadena(rs!modovtaunidadmedida)))
        modoventa.usafactor = Escadena(IIf(IsNull(rs!modovtausafactconv) Or rs!modovtausafactconv = 0, "0", "1"))
        
        MBox(1).Enabled = IIf(modoventa.modificaguia = "1" And modoventa.numeraauto = "0" And modoventa.ingpedido = "1", True, False) 'Modo de pedido
        MBox(2).Enabled = IIf(modoventa.modificaguia = "1" And modoventa.numeraauto = "0", True, False)  'Modo de factura
        MBox(3).Enabled = IIf(modoventa.modificaguia = "1" And modoventa.numeraauto = "0", True, False) 'Modo de boleta
        MBox(4).Enabled = IIf(modoventa.modificaguia = "1" And modoventa.numeraauto = "0" And modoventa.ingguia = "1", True, False) 'Modo de Modifica
        
        modoventa.numeraauto = Escadena(IIf(IsNull(rs!modovtanumautom) Or rs!modovtanumautom = 0, "0", "1"))
        
        MBox2(0).Enabled = IIf(modoventa.usafactor = 0 Or (modoventa.usafactor = 1 And modoventa.unidadmedida = "R"), True, False)
        MBox2(12).Enabled = IIf(modoventa.usafactor = 0 Or (modoventa.usafactor = 1 And modoventa.unidadmedida = "V"), True, False)
     End If
     rs.Close
     Set rs = Nothing
  End If
End Sub

Private Sub Combo3_GotFocus()
   If Combo3.ListCount - 1 <= 0 Then
      Call dllgeneral.llenacombo(Combo3, "select modovtacodigo,modovtadescripcion from vt_modoventa", VGCNx)
      Exit Sub
   End If

End Sub


Private Sub Combo3_KeyPress(KeyAscii As Integer)
  Call Combo3_Click
  Seguir Combo3, KeyAscii
End Sub

Private Sub Combo4_GotFocus()
   If Combo4.ListCount - 1 <= 0 Then
       Call dllgeneral.llenacombo(Combo4, "select formapagocodigo,formapagodescripcion from vt_formapago", VGCNx)
      Exit Sub
   End If

End Sub

Private Sub Combo4_KeyPress(KeyAscii As Integer)
   Seguir Combo4, KeyAscii
End Sub

Private Sub Combo5_KeyPress(KeyAscii As Integer)
    Seguir Combo5, KeyAscii
End Sub

Private Sub Combo6_KeyPress(KeyAscii As Integer)
    Seguir Combo6, KeyAscii
End Sub

Private Sub Combo7_KeyPress(KeyAscii As Integer)
    Seguir Combo7, KeyAscii
End Sub

Private Sub Combo8_KeyPress(KeyAscii As Integer)
    Seguir Combo8, KeyAscii
End Sub




Private Sub Ctr_Ayuda1_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
    Dim acliente As New ADODB.Recordset
    
    MBox3(0) = Trim(ColecCampos.Item(0))
    MBox3(1) = Trim(ColecCampos.Item(1))
    MBox3(2) = Trim(ColecCampos.Item(2))
    MBox3(3) = Trim(ColecCampos.Item(3))
    MBox3(4) = Trim(ColecCampos.Item(4))
    
    'Combo6.Text = dllgeneral.ComboDato(ColecCampos.Item(7))
    'Combo7.Text = dllgeneral.ComboDato(ColecCampos.Item(8))
    'Combo8.Text = dllgeneral.ComboDato(ColecCampos.Item(9))
    If IsNull(ColecCampos.Item(10)) Or Len(Trim(ColecCampos.Item(10))) = 0 Then
       text1 = numero(0)
       Text2 = numero(0)
    Else
       text1 = numero(CDbl(Trim(ColecCampos.Item(10))))
       Text2 = numero(CDbl(Trim(ColecCampos.Item(10))) * 100)
    End If
    
    lcred(0) = numero(0)
    lcred(1) = numero(0)
    
    Set acliente = VGCNx.Execute("select * from vt_cliente where clientecodigo='" & Ctr_Ayuda1.xclave & "'")
    If acliente.RecordCount > 0 Then
       Combo6.ListIndex = VerificaCombo(Combo6, acliente!clientetipopersona)
       Combo7.ListIndex = VerificaCombo(Combo7, acliente!clientetipopais)
       Combo8.ListIndex = VerificaCombo(Combo8, IIf(acliente!clientemultidireccion = 1, "S", "N"))
       lcred(0) = numero(acliente!clientesaldodolares)
       lcred(1) = numero(acliente!clientelimitecreddolar)
    End If
    acliente.Close
    Set acliente = Nothing

End Sub

Private Sub DtFechaDesde_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{tab}"
End Sub


Private Sub DtFechaHasta_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{tab}"
End Sub


Private Sub Form_Activate()
   Listado
End Sub

Private Sub Form_Load()
MostrarFormVentas Me, "C"

flag = 0
Call dllgeneral.ActivaTab(0, 1, SSTab1)
nLongicampo(1) = 1000:  nLongicampo(2) = 1200:   nLongicampo(3) = 6300:   nLongicampo(4) = 600:  nLongicampo(5) = 1200

MBox(1).Enabled = False: Label2 = ""
Call Cargacombo

Listado

MBox(13) = 0
MBox(15) = 0
MBox(18) = 0

DtFechaDesde.Value = DateAdd("m", -1, Date)
DtFechaHasta.Value = Format(Date, "dd/mm/yyyy")
   
cmdBotones(0).Picture = MDIPrincipal.ImageList2.ListImages("Nuevo").Picture
cmdBotones(1).Picture = MDIPrincipal.ImageList2.ListImages("Modificar").Picture
cmdBotones(2).Picture = MDIPrincipal.ImageList2.ListImages("Eliminar").Picture
cmdBotones(4).Picture = MDIPrincipal.ImageList2.ListImages("Retornar").Picture

cmdBotones(11).Picture = MDIPrincipal.ImageList2.ListImages("Facturar").Picture
cmdBotones(12).Picture = MDIPrincipal.ImageList2.ListImages("Retornar").Picture

End Sub

Public Function Cargacombo()
   Dim J As Integer
   Dim nsql As String
   
   CargaGrilla
   MBox2(11) = rsdeta.RecordCount
   If MBox2(11) > modoventa.nroitem Then
      MsgBox "No se puede Ingresar mas Items...Verifique!!!", vbInformation, MsgTitle
      Exit Function
   End If
  
   Call dllgeneral.llenacombo(Combo1, "select monedacodigo,monedadescripcion from gr_moneda", VGCNx)
   If Combo1.ListCount - 1 >= 0 Then
       Combo1.ListIndex = 0
   End If
   
   Combo2.Clear
   For J = 1 To 9
     Combo2.AddItem Trim(Str(J))
   Next J
   Combo2.ListIndex = 0
   
   Call dllgeneral.llenacombo(Combo3, "select modovtacodigo,modovtadescripcion from vt_modoventa", VGCNx)
   If Combo3.ListCount - 1 >= 0 Then
     Combo3.ListIndex = 0
   End If
   
   Call dllgeneral.llenacombo(Combo4, "select formapagocodigo,formapagodescripcion from vt_formapago", VGCNx)
   If Combo4.ListCount - 1 >= 0 Then
       Combo4.ListIndex = 0
   End If
   
   
   Call CargarTipoVentas(Combo5, 3)
   
   Call CargarTipoVentas(Combo6, 4)
   
   Call CargarTipoVentas(Combo7, 5)
   
   Call CargarTipoVentas(Combo8, 3)
   
   
   Call Ctr_Ayuda1.conexion(VGCNx)
   Call Ctr_Ayuda2.conexion(VGCNx)
   Call Ctr_Ayuda3.conexion(VGCNx)
   Ctr_Ayuda3.Filtro = " empresacodigo='" & VGParametros.empresacodigo & "' and puntovtacodigo='" & VGParametros.puntovta & "'"
   
   
End Function

Public Function CargaGrilla()

   Call rsdeta.Fields.Append("Item", adInteger)
   Call rsdeta.Fields.Append("Codigo", adChar, 20)
   Call rsdeta.Fields.Append("Descripcion", adChar, 100)
   Call rsdeta.Fields.Append("UM", adChar, 3)
   Call rsdeta.Fields.Append("Cant", adDouble)
   Call rsdeta.Fields.Append("Precio_Vta", adDouble)
   Call rsdeta.Fields.Append("Dscto(%)", adDouble)
   Call rsdeta.Fields.Append("Total", adDouble)
   Call rsdeta.Fields.Append("%", adDouble)
   Call rsdeta.Fields.Append("CantRef", adDouble)
   Call rsdeta.Fields.Append("Factor", adDouble)
   Call rsdeta.Fields.Append("%P", adDouble)
   
   rsdeta.Open
   If rsdeta.RecordCount > 0 Then
     Totales
   End If
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
      .Columns(2).Width = 3000
      .Columns(2).Caption = "Descripcion"
      .Columns(3).Width = 600
      .Columns(3).Caption = "UM"
      .Columns(4).Width = 1000
      .Columns(4).Caption = "Cant"
      .Columns(5).Width = 1000
      .Columns(5).Caption = "Precio_Vta"
      .Columns(6).Width = 1000
      .Columns(6).Caption = "Dscto(%)"
      .Columns(7).Width = 800
      .Columns(7).Caption = "Total"
      .Columns(8).Width = 1000
      .Columns(8).Caption = "%"
      .Columns(5).NumberFormat = "###,##0.0000"
      .Columns(6).NumberFormat = "###,##0.00"
      .Columns(7).NumberFormat = "###,##0.0000"
      .Columns(8).NumberFormat = "###,##0.00"
      .Columns(9).Width = 800
      .Columns(9).Caption = "Cant.Ref"
      .Columns(9).NumberFormat = "###,##0"
      .Columns(10).Width = 600
      .Columns(10).Caption = "Factor"
      .Columns(10).NumberFormat = "###,##0.00"
      .Columns(11).Width = 0
      .Columns(11).NumberFormat = "###,##0.00"
   End With
   TDBGrid1.Refresh
End Function
Public Function Listado()
Dim RsListado As ADODB.Recordset
Dim SQL As String

SQL = ""

If DtFechaDesde.Value > DtFechaHasta.Value Then
    MsgBox "La fecha de inicio no puede ser mayor " & Chr(13) & "a la fecha final.", vbInformation, "Sistemas"
    DtFechaDesde.SetFocus
    Exit Function
End If

If Len(Trim(TxtNro.Text)) <> 0 Then SQL = " and pedidonumero='" & TxtNro.Text & "'"
If Len(Trim(TxtCliente.Text)) <> 0 Then SQL = SQL + " and clienterazonsocial like '%" & TxtCliente.Text & "%' "

Set RsListado = VGCNx.Execute("select pedidonumero as Pedido,pedidofecha as Fecha,clienterazonsocial as Cliente,pedidonotaped as Cotizacion," _
& " pedidomensaje as Descripcion " _
& " from cotizalibre where empresacodigo='" & VGParametros.empresacodigo & "' and puntovtacodigo='" & VGParametros.puntovta & "' " _
& " and pedidofecha between '" & DtFechaDesde.Value & "' and '" & DtFechaHasta.Value & "' " & SQL & " " _
& " order by pedidofecha,pedidonumero")

LblReg.Caption = "(" & RsListado.RecordCount & ") Cotizaciones"
TDBGrid2.DataSource = RsListado

With TDBGrid2
  .Columns(0).Width = 1200
  .Columns(1).Width = 1200
  .Columns(2).Width = 3300
  .Columns(3).Width = 2000
  .Columns(4).Width = 6500
  .AllowUpdate = False
  .Refresh
End With

End Function

Private Sub Form_Unload(Cancel As Integer)
  Set rsdeta = Nothing
End Sub

Private Sub MBox_GotFocus(Index As Integer)
'Call dllgeneral.Enfoquetexto(MBox(Index))
End Sub

Private Sub MBox_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 And Index >= 5 And Index < 19 Then
    If Index = 9 Then
      SSTab2.Tab = 1
      Combo3.SetFocus
    Else
      If Index Like "[567]" Then
         Totales
      End If
      SendKeys "{tab}"
    End If
  ElseIf KeyCode = 13 And (Index = 19 Or Len(Trim(MBox(Index))) > 0) Then
'    MBox2(0).SetFocus
    SendKeys "{tab}"
    Exit Sub
  End If
End Sub

Private Sub MBox_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = 13 And Index = 19 Then
  '   MBox2(0).SetFocus
  '   Exit Sub
  End If
End Sub

Private Sub MBox_LostFocus(Index As Integer)
  On Error Resume Next
  Select Case Index
   Case 5, 6, 7, 8, 13, 15
      If Not dllgeneral.ValidaCadena(MBox(Index), "N") Then
'         MsgBox Msg29, vbInformation, "AVISO"
'         Call dllgeneral.Enfoquetexto(MBox(Index))
        'If Index = 13 Then
        MBox(Index) = 0
        ' Exit Sub
      End If
      MBox(Index) = Format(MBox(Index), "##,##0.00")
   Case 10
      If Not dllgeneral.ValidaCadena(MBox(Index), "F") Then
'         MsgBox "Fecha No Valida", vbInformation, "AVISO"
'         Call dllgeneral.Enfoquetexto(MBox(Index))
         Exit Sub
      End If
   Case 16, 17
      If Not dllgeneral.ValidaCadena(MBox(Index), "D") Then
 '        MsgBox Msg29, vbInformation, "AVISO"
         'Call dllgeneral.Enfoquetexto(MBox(Index))
         Exit Sub
      End If
      MBox(Index) = Right("000000000000" & MBox(Index), MBox(Index).MaxLength)
   Case 18
      If Not dllgeneral.ValidaCadena(MBox(Index), "D") Then
         MsgBox Msg29, vbInformation, "AVISO"
         Call dllgeneral.Enfoquetexto(MBox(Index))
         Exit Sub
      End If
      MBox(Index) = Format(MBox(Index), "####0")
      Exit Sub
  Case 19
        If IsNull(MBox(19)) Or Len(Trim(MBox(19))) = 0 Then
           'MsgBox "Falta Punto de LLegada", vbInformation, MsgTitle
           'Call dllgeneral.Enfoquetexto(MBox(19))
           Exit Sub
        End If
        If IsNull(Ctr_Ayuda1.xclave) Or Len(Trim(Ctr_Ayuda1.xclave)) = 0 Then
           MsgBox W1TXT1, vbInformation, MsgTitle
           SSTab2.Tab = 1
           Ctr_Ayuda1.SetFocus
           Exit Sub
        End If
        If IsNull(Ctr_Ayuda2.xclave) Or Len(Trim(Ctr_Ayuda2.xclave)) = 0 Then
           MsgBox W1TXT6, vbInformation, MsgTitle
           SSTab2.Tab = 1
           Ctr_Ayuda2.SetFocus
           Exit Sub
        End If
        If IsNull(Ctr_Ayuda3.xclave) Or Len(Trim(Ctr_Ayuda3.xclave)) = 0 Then
           MsgBox W1TXT7, vbInformation, MsgTitle
           SSTab2.Tab = 1
           Ctr_Ayuda3.SetFocus
           Exit Sub
        End If
        If IsNull(MBox(8)) Or Len(Trim(MBox(8))) = 0 Or CDbl(MBox(8)) <= 0 Then
           MsgBox "Falta Tipo de Cambio", vbInformation, MsgTitle
           SSTab2.Tab = 0
           Call dllgeneral.Enfoquetexto(MBox(8))
           Exit Sub
        End If
        If IsNull(MBox(15)) Or Len(Trim(MBox(15))) = 0 Or CDbl(MBox(15)) < 0 Then
           MsgBox W1TXT9, vbInformation, MsgTitle
           SSTab2.Tab = 1
           Exit Sub
        End If
        If IsNull(MBox2(10)) Or Len(Trim(MBox2(10).ClipText)) = 0 Then
           MsgBox W1TXT30, vbInformation, MsgTitle
           Exit Sub
        End If
        
      Fr1(0).Enabled = False
      Fr2(0).Enabled = False
      Fr3(0).Enabled = False
      TClie.Enabled = False
      Call Combo3_Click
        
     ' MBox2(0).SetFocus
      Exit Sub
   Case 9
      Call MBox_KeyDown(9, 13, 0)
      Exit Sub
      
   Case 2, 3, 4
        MBox(Index) = Right("000000000000" & MBox(Index), MBox(Index).MaxLength)
  End Select
End Sub

Private Sub MBox2_GotFocus(Index As Integer)
  On Error Resume Next
  If Index = 3 Then 'And dllgeneral.ComboDato(Combo5) = "N" Then
     Call TraerProducto
  End If
  If Index Like "[234]" Then
        Fr1(0).Enabled = False
        Fr2(0).Enabled = False
        Fr3(0).Enabled = False
        TClie.Enabled = False
   End If
  Call dllgeneral.Enfoquetexto(MBox2(Index))
End Sub

Private Sub MBox2_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   On Error Resume Next
  If KeyCode = 13 Then
    If Index = 12 Then
      MBox2(Index) = Format(MBox2(Index), "##,###,##0")
    End If
   If Index = 1 Then
     Call TraerProducto
   End If
   SendKeys "{tab}"
  ElseIf Index = 1 Then
      If dllgeneral.ValidaCadena(Trim(MBox2(1).ClipText), "N") = False Then
        MBox2(1).MaxLength = 64
      Else
        MBox2(1).MaxLength = 8
      End If
  End If
End Sub

Private Sub MBox2_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Index Like "[345]" Or Index Like "[01]" Then
'          If Index = 1 And Len(Trim(MBox2(0).ClipText)) = 0 Then
'             MsgBox W1TXT22, vbInformation, MsgTitle
'             MBox2(0).SetFocus
'             Exit Sub
'          End If
        '  SendKeys "{tab}"
      End If
   End If
End Sub

Private Sub MBox2_LostFocus(Index As Integer)
  
  Dim nregi As Long
  Dim wposi, posi As Integer
  Dim ntabla As String
  Dim wflag As Integer
  
  On Error Resume Next
  
  Select Case Index
   Case 0
      If Not dllgeneral.ValidaCadena(MBox2(Index), "N") And Len(Trim(MBox2(Index))) > 0 Then
         MsgBox Msg29, vbInformation, "AVISO"
         Call dllgeneral.Enfoquetexto(MBox2(Index))
         Exit Sub
      End If
   Case 1
      If dllgeneral.VerificaDatoExistente(VGCNx, "select * from stkart where stcodigo='" & MBox2(Index).Text & "' and stalma='" & Ctr_Ayuda3.xclave & "' ") = 0 And Len(Trim(MBox2(Index))) > 0 Then
         Call cAyuda_Click(3)
         MBox2(1).MaxLength = 8
         Exit Sub
      Else
        wflag = verificaproducto()
        If wflag = 1 Then
            MsgBox "Ya ingreso el producto...Verifique!!!", vbInformation, MsgTitle
            MBox2(1).SetFocus
            Exit Sub
         End If
            
      End If
   Case 3, 4, 5
      If Index = 3 And dllgeneral.ComboDato(Combo5) = "N" Then
         Call TraerProducto
      End If
      If Not dllgeneral.ValidaCadena(MBox2(Index), "N") And Len(Trim(MBox2(Index))) > 0 Then
         MsgBox Msg29, vbInformation, "AVISO"
         Call dllgeneral.Enfoquetexto(MBox2(Index))
         Exit Sub
      End If
      If Not dllgeneral.ValidaCadena(MBox2(0), "N") Then
         MsgBox Msg29, vbInformation, "AVISO"
         Call dllgeneral.Enfoquetexto(MBox2(0))
         Exit Sub
      End If
      wflag = verificaproducto()
      If wflag = 1 Then
        Label2 = ""
        MsgBox "Ya ingreso el producto...Verifique!!!", vbInformation, MsgTitle
        MBox2(1).SetFocus
        Exit Sub
      End If
      If Index = 5 Then
         If Trim(MBox2(3)) = "" Or Trim(MBox2(4)) = "" Or Trim(MBox2(5)) = "" Then
           MsgBox Msg29, vbInformation, "AVISO"
           Call dllgeneral.Enfoquetexto(MBox2(1))
           Exit Sub
         End If
      End If
      If Index Like "[45]" Then
         MBox2(Index) = numero(MBox2(Index))
       Else
         MBox2(Index) = Format(MBox2(Index), "######0.000")
       End If
       If Index = 5 And Len(Trim(MBox2(Index))) > 0 Then
        If modoventa.nroitem < TDBGrid1.ApproxCount Then
           MsgBox "Excede el Numero de Items del Documento..!!", vbInformation, MsgTitle
           Exit Sub
        End If
        nregi = 0
        wposi = 0
        If rsdeta.RecordCount > 0 Then
            rsdeta.MoveLast
            wposi = rsdeta.Fields(0)
            posi = rsdeta.Fields(0)
            rsdeta.MoveFirst
            Do Until rsdeta.EOF
               If rsdeta.Fields(0) = MBox2(11) Then
                  posi = rsdeta.Fields(0)
                  Exit Do
               End If
               nregi = nregi + 1
               rsdeta.MoveNext
            Loop
        End If
        If rsdeta.RecordCount = nregi Then
          wposi = wposi + 1
          posi = wposi
          rsdeta.AddNew
        End If
        rsdeta.Fields(0) = posi
        rsdeta.Fields(1) = Escadena(MBox2(1))
        rsdeta.Fields(2) = Left(Escadena(Label2) & Space(40), 40)
        rsdeta.Fields(3) = Trim(MBox2(2))
        rsdeta.Fields(4) = Escadena(MBox2(0))
        If VGParamSistem.tieneigv = "1" Then
           rsdeta.Fields(5) = MBox2(3)          '(MBox2(3) / (1 + VGParamSistem.Igv/100))
        Else
           If modoventa.impuestos = "1" Then
              rsdeta.Fields(5) = (MBox2(3) / (1 + VGParamSistem.Igv / 100))
           Else
              rsdeta.Fields(5) = MBox2(3)
           End If
        End If
        rsdeta.Fields(6) = numero(MBox2(4))
        rsdeta.Fields(7) = numero(MBox2(0) * MBox2(3))   ' IIf(VGParamSistem.tieneigv = "1", (MBox2(3) / (1 + (VGParamSistem.igv / 100))), MBox2(3)))
        rsdeta.Fields(8) = numero(MBox2(5))
        rsdeta.Fields(9) = IIf(Len(Trim(MBox2(11))) = 0, 0, Format(MBox2(11), "##,###,##0"))
        rsdeta.Fields(10) = numero(MBox2(13))
        rsdeta.Fields(11) = numero(IIf(IsNull(MBox2(14)) Or Len(Trim(MBox2(14))) = 0, 0, MBox2(14)))
        rsdeta.Update
        TDBGrid1.Row = rsdeta.RecordCount - 1
        
        ConfigGrid
        Totales
        MBox2(11) = wposi + 1
        If MBox2(12).Enabled = True Then
          MBox2(12).SetFocus
        Else
          MBox2(0).SetFocus
        End If
        flag = 0
        Exit Sub
    End If
  End Select

End Sub

Private Sub MBox3_KeyPress(Index As Integer, KeyAscii As Integer)
   Seguir MBox3(Index), KeyAscii
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
  If SSTab1.Tab = 2 Then
     MBox2(0).SetFocus
  ElseIf SSTab1.Tab = 1 Then
     If MBox(0).Enabled = True Then
        MBox(5).SetFocus
     Else
        If MBox(5).Enabled = True Then MBox(5).SetFocus
     End If
  End If
End Sub

 
 
Public Function Totales()
Dim J As Double
Dim Previo As Double
Dim rssql As New ADODB.Recordset
Dim SQL As String
Dim dct01, dct02, dct03, dct04, dct05, dct06 As Double
  
Tbruto = 0: Tigv = 0: Tdscto = 0: TNeto = 0: TCant = 0
TImporte = 0: TSub = 0
'--Totales de Descuentos
DTGlobal = 0: DTCliente = 0: DTPPago = 0: DTOficina = 0: DTItem = 0
DTLinea = 0: DTPromo = 0: MBox2(6) = 0

   
  If rsdeta.RecordCount > 0 Then
    rsdeta.MoveFirst
    For J = 0 To rsdeta.RecordCount - 1
       'IMPORTE DE MONTO BRUTO SIN IGV, ES DECIR PRECIO X CANTIDAD
'       Tbruto = Tbruto + ((rsdeta.Fields(7) - (rsdeta.Fields(7) * MBox(5) / 100)) / (1 + VGParamSistem.Igv/100))
        Tbruto = Tbruto + ((rsdeta.Fields(7) - (rsdeta.Fields(7) * MBox(5) / 100)))
        
       TCant = TCant + rsdeta.Fields(4)
       TImporte = rsdeta.Fields(4) * rsdeta.Fields(5)           '(rsdeta.Fields(7) + rsdeta.Fields(7))  'rsdeta.Fields(4) *
       
       If IsNull(text1) Or Len(Trim(text1)) = 0 Then
           dct06 = 0
       Else
           dct06 = 0
       End If
       
        dct01 = 0    ' descuento por cliente
        DTCliente = DTCliente + dct01
        
        'DESCUENTO POR ITEM
        dct02 = 0
        dct02 = (TImporte * (rsdeta.Fields(6) / 100))
        
        DTItem = DTItem + dct02
        
        'DESCUENTO ESPECIAL  :w8dct03 =(w8bruto - w8dct02-w8dct06)*w2dctpp/100
         dct03 = (TImporte - dct02 - dct06) * (MBox(7) / 100)
         
        DTPPago = DTPPago + dct03
         
        'DESCUENTO POR PROMOCION  : w8dct04 =(w8bruto - w8dct02-w8dct03-w8dct06)*w2dctpr/100
        dct04 = (TImporte - dct02 - dct03 - dct06) * (MBox(6) / 100)
        DTPromo = DTPromo + dct04
         
        'DESCUENTO GENERAL : w8dct05 =(w8bruto - w8dct02-w8dct03-w8dct04-w8dct06)*w2dctgl/100
        dct05 = (TImporte - dct02 - dct03 - dct04 - dct06) * (MBox(5) / 100)
                 
        DTGlobal = DTGlobal + dct05
        
        'ACUMULADO DE TOTAL DESCUENTOS  :w8dctos = w8dct02 + w8dct03+w8dct04+w8dct05+w8dct06
         Tdscto = Tdscto + (dct01 + dct02 + dct03 + dct04 + dct05 + dct06)
    
       'ACUMULADO DE SUBTOTAL DE VENTA : w8subto = w8bruto - w8dctos
        TSub = TSub + (TImporte - (dct01 + dct02 + dct03 + dct04 + dct05 + dct06))
                
'       If VGParamSistem.tieneigv = "1" Then
'            Previo = (TImporte - (dct01 + dct02 + dct03 + dct04 + dct05 + dct06))
'            Previo = Previo - (Previo / (1 + VGParamSistem.Igv/100)) '(Previo * VGParamSistem.Igv)
'            Tigv = Tigv + Previo
'       Else
'           If modoventa.impuestos = "1" Then
'                SQL = " select tieneigv from  grupo,maeart where acodigo='" & rsdeta.Fields(1) & "'"
'                SQL = SQL & " and afamilia=fam_codigo and alinea=lin_codigo and agrupo=gru_codigo "
'                Set rssql = VGCNx.Execute(SQL)
'                If rssql.RecordCount > 0 Then
'                   If rssql!tieneigv = "1" Then
'                      Previo = (TImporte - (dct01 + dct02 + dct03 + dct04 + dct05 + dct06))
'                      Previo = (Previo * VGParamSistem.Igv)
'                      Tigv = Tigv + Previo
'                      Previo = (TImporte - (dct01 + dct02 + dct03 + dct04 + dct05 + dct06))
'                      Previo = (Previo * (1 + VGParamSistem.Igv/100))
'                      rsdeta.Fields(7) = Previo
'                    Else
'                      Previo = (TImporte - (dct01 + dct02 + dct03 + dct04 + dct05 + dct06))
'                      Previo = (Previo * 0)
'                      Tigv = Tigv + Previo
'                      Previo = (TImporte - (dct01 + dct02 + dct03 + dct04 + dct05 + dct06))
'                      Previo = (Previo * (1 + 0))
'                      rsdeta.Fields(7) = Previo
'
'                   End If
'                 Else
'                   Previo = (TImporte - (dct01 + dct02 + dct03 + dct04 + dct05 + dct06))
'                   Previo = (Previo * 0)
'                   Tigv = Tigv + Previo
'
'                   Previo = (TImporte - (dct01 + dct02 + dct03 + dct04 + dct05 + dct06))
'                   Previo = (Previo * (1 + 0))
'                   rsdeta.Fields(7) = Previo
'               End If
'           Else
'               If rsdeta.Fields(11) > 0 Then
'                    Previo = (TImporte - (dct01 + dct02 + dct03 + dct04 + dct05 + dct06))
'                    rsdeta.Fields(7) = Previo * (1 + rsdeta(11))
'                    Tigv = Tigv + (Previo * rsdeta(11))
'               Else
'                    Previo = (TImporte - (dct01 + dct02 + dct03 + dct04 + dct05 + dct06))
'                    rsdeta.Fields(7) = Previo
'                    'Tigv = Tigv
'              End If
'           End If
'        End If
        rsdeta.Update
      
       rsdeta.MoveNext
    Next J
  Else
    Exit Function
  End If
  
 'IMPORTE TOTAL NETO DE FACTURA   w8tneto = w8subto + w8impto
  TNeto = Tbruto + (Tbruto * VGParamSistem.Igv)     'Tbruto - Tdscto + Tigv
  MBox2(7) = Format(Tbruto, "#,###,##0.0000")
  MBox2(6) = numero(TCant)
  MBox2(9) = numero(Tbruto * VGParamSistem.Igv) 'numero(Tigv)
  MBox2(8) = numero(Tdscto)
  MBox2(10) = numero(TNeto)
  
'  TNeto = Tbruto + Tigv      'Tbruto - Tdscto + Tigv
'  MBox2(7) = Format(Tbruto, "#,###,##0.0000")
'  MBox2(6) = numero(TCant)
'  MBox2(9) = numero(Tigv)
'  MBox2(8) = numero(Tdscto)
'  MBox2(10) = numero(TNeto)
  
  Limpiartexto MBox2, 12, 12
  Limpiartexto MBox2, 13, 13
  Limpiartexto MBox2, 14, 14
  Limpiartexto MBox2, 0, 5
'**********************************************************************************
'  Dim j As Double
'  Dim Previo As Double
'  Dim dct02, dct03, dct04, dct05, dct06 As Double
'
'  Tbruto = 0: Tigv = 0: Tdscto = 0: TNeto = 0: TCant = 0
'  TImporte = 0: TSub = 0
'  '--Totales de Descuentos
'  DTGlobal = 0: DTCliente = 0: DTPPago = 0: DTOficina = 0: DTItem = 0
'  DTLinea = 0: DTPromo = 0
'
'  If rsdeta.RecordCount > 0 Then
'    rsdeta.MoveFirst
'    For j = 0 To rsdeta.RecordCount - 1
'       'IMPORTE DE MONTO BRUTO SIN IGV, ES DECIR PRECIO X CANTIDAD
'
'       Tbruto = Tbruto + (rsdeta.Fields(4) * rsdeta.Fields(5))
'       TCant = TCant + rsdeta.Fields(4)
'       TImporte = (rsdeta.Fields(4) * rsdeta.Fields(5))
'
'       'DESCUENTO DE CIA O EMPRESA
''       If VGParamSistem.tienedscto = "1" Then
''            dct06 = TImporte * (1 + VGParamSistem.descuento)
''       Else
'''           dct06 = 0
''       End If
'       If IsNull(Text1) Or Len(Trim(Text1)) = 0 Then
'           dct06 = 0
'       Else
'          'dct06 = TImporte * (1 + VGParamSistem.descuento)
'          dct06 = TImporte * (CDbl(Text1))
'       End If
'
'       DTCliente = DTCliente + dct06
'
'       'DESCUENTO POR ITEM
'       dct02 = 0
'       dct02 = (TImporte * (rsdeta.Fields(6) / 100))
'
'       DTItem = DTItem + dct02
'
'       'DESCUENTO ESPECIAL  :w8dct03 =(w8bruto - w8dct02-w8dct06)*w2dctpp/100
'        dct03 = (TImporte - dct02 - dct06) * (MBox(7) / 100)            '(Tbruto-dct02-dct06)
'
'        DTPPago = DTPPago + dct03
'
'       'DESCUENTO POR PROMOCION  : w8dct04 =(w8bruto - w8dct02-w8dct03-w8dct06)*w2dctpr/100
'        dct04 = (TImporte - dct02 - dct03 - dct06) * (MBox(6) / 100)
'
'        DTPromo = DTPromo + dct04
'
'       'DESCUENTO GENERAL : w8dct05 =(w8bruto - w8dct02-w8dct03-w8dct04-w8dct06)*w2dctgl/100
'        dct05 = (TImporte - dct02 - dct03 - dct04 - dct06) * (MBox(5) / 100)
'        DTGlobal = DTGlobal + dct05
'
'       'ACUMULADO DE TOTAL DESCUENTOS  :w8dctos = w8dct02 + w8dct03+w8dct04+w8dct05+w8dct06
'        Tdscto = Tdscto + (dct02 + dct03 + dct04 + dct05 + dct06)
'
'       'ACUMULADO DE SUBTOTAL DE VENTA : w8subto = w8bruto - w8dctos
'        TSub = TSub + (TImporte - (dct02 + dct03 + dct04 + dct05 + dct06))
'
'       If VGParamSistem.tieneigv = "1" Then
'            'CALCULAMOS EL IMPORTE :=  TOTAL IMPORTE SIN IGV - DESCTOS + IGV
'            Previo = (TImporte - (dct02 + dct03 + dct04 + dct05 + dct06))
'            Previo = (Previo * VGParamSistem.Igv)
'            Tigv = Tigv + Previo
'
'            'GRABAMOS EL TOTAL DE IMPORTE EN LA TABLA TEMPORAL PARA MOSTRAR
'            Previo = (TImporte - (dct02 + dct03 + dct04 + dct05 + dct06))
'            Previo = (Previo * (1 + VGParamSistem.Igv/100))
'            rsdeta.Fields(7) = Previo
'       Else                    'If VGParamSistem.tieneigv = "0" Then
'           If modoventa.impuestos = "1" Then
'                Previo = (TImporte - (dct02 + dct03 + dct04 + dct05 + dct06))
'                Previo = (Previo * VGParamSistem.Igv)
'                Tigv = Tigv + Previo
'
'                Previo = (TImporte - (dct02 + dct03 + dct04 + dct05 + dct06))
'                Previo = (Previo * (1 + VGParamSistem.Igv/100))
'                rsdeta.Fields(7) = Previo
'           Else
'               If rsdeta(11) > 0 Then
'                    Previo = (TImporte - (dct02 + dct03 + dct04 + dct05 + dct06))
'                    rsdeta.Fields(7) = Previo * (rsdeta(11))
'                    Tigv = Tigv + (Previo * rsdeta(11))
'               Else
'                    Previo = (TImporte - (dct02 + dct03 + dct04 + dct05 + dct06))
'                    rsdeta.Fields(7) = Previo
'                    Tigv = Tigv
'              End If
'           End If
'       End If
'       rsdeta.Update
'
'       rsdeta.MoveNext
'    Next j
'  Else
'    Exit Function
'  End If
'
'  ' w2imp = IIf(w2ciaimp, w2timp, pro_pctimp)
'  ' w2imp = IIf(vtmod.mod_imp, w2imp, 0)
'  ' w2prepac = IIf(w2dctofe > 0, roun(w2prepac * (100 - w2dctofe) / 100, 4), w2prepac
'
'   'set deci to 12
'   'w8bruto = w2cant  * w2prepac                                          && Total Bruto
'   'w2dctofe = 0
'   'If w2fchatn>=pro_fchini and w2fchatn<=pro_fchfin                      && Precio de Oferta en una lista de precios
'   '   w2dctofe =pro_dctofi      && Descuentos Ofertas
'   '*   w8dct01 = w2cant  * Abs(IIF(w2prelis>w2prepac,w2prelis-w2prepac,0))   && Dcto.Oferta
'   'w8dct06 = w8bruto * w0dcto/100                                        && Dcto. por Default
'   'w8dct02 = (w8bruto-w8dct06)*w2dctlin/100                              && Dcto.Por Item
'
'  'IMPORTE TOTAL NETO DE FACTURA   w8tneto = w8subto + w8impto
'  TNeto = Tbruto - Tdscto + Tigv
'  MBox2(7) = Format(Tbruto, "#,###,##0.0000")
'  MBox2(6) = numero(TCant)
'  MBox2(9) = numero(Tigv)
'  MBox2(8) = numero(Tdscto)
'  MBox2(10) = numero(TNeto)
'  Limpiartexto MBox2, 0, 5
'  Limpiartexto MBox2, 12, 12
'  Limpiartexto MBox2, 13, 13
'  Limpiartexto MBox2, 14, 14
  
End Function

Private Sub tclie_Click()
       
   SSTab2.TabEnabled(2) = IIf(TClie.Value = 1, 1, 0)
   If TClie.Value = 1 Then
        SSTab2.Tab = 2
        MBox3(0) = g_Eventual
        MBox3(0).Enabled = False
        MBox3(1).SetFocus
   End If
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
        Limpiartexto MBox2, 6, 10
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
     Totales
     Exit Sub
  ElseIf KeyCode = 13 Then
    Limpiartexto MBox2, 0, 5
    MBox2(11) = TDBGrid1.Columns(0).Text
    MBox2(0) = TDBGrid1.Columns(4).Text
    MBox2(1) = TDBGrid1.Columns(1).Text
    Label2 = TDBGrid1.Columns(2).Text
    MBox2(2) = Escadena(TDBGrid1.Columns(3).Text)
    MBox2(12) = Escadena(TDBGrid1.Columns(9).Text)
    MBox2(13) = Escadena(TDBGrid1.Columns(10).Text)
    MBox2(14) = Escadena(TDBGrid1.Columns(11).Text)
    
    If VGParamSistem.tieneigv = "1" Then
       MBox2(3) = Format(TDBGrid1.Columns(5).Text * (1 + (VGParamSistem.Igv)), "######0.000")
    Else
       If modoventa.impuestos = "1" Then
           MBox2(3) = Format(IIf(IsNull(TDBGrid1.Columns(5).Text) Or Len(Trim(TDBGrid1.Columns(5).Text)) = 0, 0, TDBGrid1.Columns(5).Text) * (1 + (VGParamSistem.Igv)), "######0.000")
       Else
           MBox2(3) = Format(TDBGrid1.Columns(5).Text, "######0.000")
       End If
    End If
    MBox2(4) = numero(TDBGrid1.Columns(6).Text)
    MBox2(5) = numero(TDBGrid1.Columns(8).Text)
    If MBox2(12).Enabled = True Then
      MBox2(12).SetFocus
    Else
      MBox2(0).SetFocus
    End If
    flag = 1
  End If
  
End Sub




Public Function Carga_Pedido()
    Dim csql As New ADODB.Recordset
    Dim acliente As New ADODB.Recordset
    Dim J As Integer
    Set csql = VGCNx.Execute("select * from cotizalibre where pedidonumero='" & TDBGrid2.Columns(0).Text & "'")
    If csql.RecordCount > 0 Then
       MBox(0) = Escadena(csql!puntovtacodigo)                    'Pto Venta
       MBox(1) = Escadena(csql!pedidonumero)                      'nro pedido
      If Escadena(csql!pedidotipofac) = g_tipofac Then
         MBox(2) = Escadena(csql!pedidonrofact)                     'nro factura
       Else
         MBox(2) = 0
       End If
       If Escadena(csql!pedidotipofac) = g_tipobol Then
          MBox(3) = Escadena(csql!pedidonrofact)                   'nro boleta
       Else
          MBox(3) = 0
       End If
       If Escadena(csql!pedidotipofac) = g_tipoguia Then
            MBox(4) = Escadena(csql!pedidonrogiarem)                   'nro guia
       Else
            MBox(4) = 0
       End If
       MBox(5) = numero(csql!pedidodsctoglobal)                   'dscto gral
       MBox(6) = numero(csql!pedidodsctoppago)                    'dscto promocional
       MBox(7) = numero(csql!pedidodsctovtaoficina)               'dscto especial
       Combo1.ListIndex = VerificaCombo(Combo1, csql!pedidomoneda)     'moneda
       MBox(8) = numero(csql!pedidotipcambio)                             'tipo de cambio
       Combo2.ListIndex = VerificaCombo(Combo2, Trim(csql!pedidolistaprec)) 'lista precios
       TxtContacto.Text = Escadena(csql!pedidomensaje)                            'mensajes
       Combo3.ListIndex = VerificaCombo(Combo3, csql!modovtacodigo)       'modo de venta
       MBox(10) = Format(csql!pedidofecha, "dd/mm/yyyy")                            'fecha de atencion
       Combo4.ListIndex = VerificaCombo(Combo4, csql!formapagocodigo) 'forma de pago
       Ctr_Ayuda1.xclave = Escadena(csql!clientecodigo)                  ' cliente MBox(11)
       
       '*****Respecto a Clientes *******
       Call Ctr_Ayuda1.Ejecutar
       Set acliente = VGCNx.Execute("select * from vt_cliente where clientecodigo='" & Ctr_Ayuda1.xclave & "'")
       If acliente.RecordCount > 0 Then
          Combo6.ListIndex = VerificaCombo(Combo6, acliente!clientetipopersona)
          Combo7.ListIndex = VerificaCombo(Combo7, acliente!clientetipopais)
          Combo8.ListIndex = VerificaCombo(Combo8, IIf(acliente!clientemultidireccion = 1, "S", "N"))
          lcred(0) = numero(acliente!clientesaldodolares)
          lcred(1) = numero(acliente!clientelimitecreddolar)
       End If
       acliente.Close
       Set acliente = Nothing
       
       Ctr_Ayuda2.xclave = Escadena(csql!vendedorcodigo)                    'vendedor
       Call Ctr_Ayuda2.Ejecutar
       MBox(13) = numero(csql!pedidoporccomision)                           'comision
       Ctr_Ayuda3.xclave = Escadena(csql!almacencodigo)                     'almacen
       Call Ctr_Ayuda3.Ejecutar
       MBox(15) = numero(csql!pedidototalotros)                             'otros gastos
       MBox(16) = Escadena(csql!pedidonotaped)                              'nota pedido
       MBox(17) = Escadena(csql!pedidoordencompra)                          'orden de compra
       Combo5.ListIndex = VerificaCombo(Combo5, IIf(csql!pedidoautorizacion = 1, "S", "N")) 'autorizacion
       MBox(18) = Format(csql!pedidodiaspago, "##0")                        'dias pago
       MBox2(6) = numero(csql!pedidototitem)                                'Total Cantidad
       MBox2(7) = numero(csql!pedidototbruto)                               'Total Bruto
       MBox2(8) = numero(csql!pedidototalflete)                             'Total Dsctos
       MBox2(9) = numero(csql!pedidototimpuesto)                            'Total Igv
       MBox2(10) = numero(csql!pedidototneto)                               'Neto a Facturar
       MBox(19) = Escadena(csql!pedidoentrega)                             'Entrega de Pedidos
       TClie.Value = 0
       SSTab2.Tab = 0
       SSTab2.TabEnabled(2) = True
    End If
    csql.Close
       
                          
    Set csql = VGCNx.Execute("select detpeditem,A.productocodigo,b.adescri,a.unidadcodigo," & _
                          "detpedcantpedida,detpedmontoprecvta,detpeddsctoxitem,detpedimpbruto," & _
                          " detpedporccomis,detpedcantpedidaref,detpedfactorconv " & _
                          "from detallecotizalibre A " & _
                          "inner Join " & _
                          "[" & VGCNx.DefaultDatabase & "].dbo.maeart B" & _
                          " ON A.productocodigo=b.acodigo COLLATE Modern_Spanish_CI_AI " & _
                          " where pedidonumero='" & TDBGrid2.Columns(0).Text & "' ")
                          'and B.almacencodigo='" & Ctr_Ayuda3.xclave & "'
    Set rsdeta = Nothing
    Call CargaGrilla
   
    Do Until csql.EOF
       rsdeta.AddNew
       rsdeta.Fields(0) = Escadena(csql!detpeditem)
       rsdeta.Fields(1) = Escadena(csql!productocodigo)
       rsdeta.Fields(2) = Escadena(csql!adescri)
       rsdeta.Fields(3) = Escadena(csql!unidadcodigo)
       rsdeta.Fields(4) = numero(csql!detpedcantpedida)
       rsdeta.Fields(5) = numero(IIf(IsNull(csql!detpedmontoprecvta), 0, csql!detpedmontoprecvta))
       rsdeta.Fields(6) = numero(csql!detpeddsctoxitem)
       rsdeta.Fields(7) = numero(csql!detpedimpbruto)
       rsdeta.Fields(8) = numero(csql!detpedporccomis)
       rsdeta.Fields(9) = numero(csql!detpedcantpedidaref)
       rsdeta.Fields(10) = numero(IIf(IsNull(csql!detpedfactorconv), 0, csql!detpedfactorconv))
       rsdeta.Update
       csql.MoveNext
    Loop
    csql.Close
    
    Call ConfigGrid
    Set csql = Nothing

End Function



Public Function GrabarData() As Integer
    Dim J As Integer
    Dim regi As Long
    Dim nsql As String
    Dim ltipo As String
    Dim Previo As Double
    Dim dct02, dct03, dct04, dct05, dct06 As Double
    Dim tinafecto As Double
    Dim acmd As New ADODB.Command
    Dim asql As New ADODB.Recordset
    Dim valor As Double
    On Error GoTo vererror
    
    
    GrabarData = 0
    
    '******** CABECERA DE MOVIMIENTO *****************
    If rsdeta.RecordCount = 0 Then
      MsgBox W1TXT30, vbInformation, MsgTitle
      GrabarData = 0
      Exit Function
    End If
    'Call Totales
    For J = 1 To 29
        wCabe(J) = ""
    Next J
    valor = 0
    If TDBGrid2.ApproxCount > 0 Then
          TDBGrid2.MoveFirst
          Do Until TDBGrid2.EOF
              valor = CDbl(TDBGrid2.Columns(0).Text)
              TDBGrid2.MoveNext
         Loop
    Else
        valor = 1
    End If
    'MBox(1) = Right("000000000000" & Trim(CStr(valor)), MBox(1).MaxLength)
    wCabe(1) = MBox(0)                       'Pto Venta
    wCabe(2) = Trim(MBox(1))                       'nro pedido
    wCabe(3) = Trim(MBox(2))                        'nro factura
    wCabe(4) = Trim(MBox(3))                         'nro boleta
    wCabe(5) = Trim(MBox(4))                         'nro guia
    wCabe(6) = MBox(5)                       'dscto gral
    wCabe(7) = MBox(6)                       'dscto promocional
    wCabe(8) = MBox(7)                       'dscto especial
    wCabe(9) = dllgeneral.ComboDato(Combo1.Text)        'moneda
    wCabe(10) = MBox(8)                      'tipo de cambio
    wCabe(11) = dllgeneral.ComboDato(Combo2.Text)       'lista de precios
    wCabe(12) = TxtContacto.Text                      'mensajes
    wCabe(13) = dllgeneral.ComboDato(Combo3.Text)       'modo de venta
    wCabe(14) = MBox(10)                     'fecha de atencion
    wCabe(15) = dllgeneral.ComboDato(Combo4.Text)       'forma de pago
    wCabe(16) = MBox3(0)    'Ctr_Ayuda1.xclave         ' MBox(11)                     'cliente
    wCabe(17) = Ctr_Ayuda2.xclave        'MBox(12)                     'vendedor
    wCabe(18) = MBox(13)                  'comision
    wCabe(19) = Ctr_Ayuda3.xclave        'MBox(14)                     'almacen
    wCabe(20) = MBox(15)                     'otros gastos
    wCabe(21) = MBox(16)                     'nota pedido
    wCabe(22) = MBox(17)                     'orden de compra
    wCabe(23) = dllgeneral.ComboDato(Combo5.Text)       'autorizacion
    wCabe(24) = MBox(18)                     'dias pago
    wCabe(25) = MBox2(6)                    'Total Cantidad
    wCabe(26) = MBox2(7)                    'Total Bruto
    wCabe(27) = 0    'MBox2(8)              'Total Dsctos   -- total fletes
    wCabe(28) = MBox2(9)                    'Total Igv
    wCabe(29) = MBox2(10)                   'Neto a Facturar
    wCabe(30) = MBox(19)                    'entrega pedido
    wCabe(31) = MBox3(1)                    'nombre cliente
    wCabe(32) = MBox3(3)                    'direccion
    wCabe(33) = MBox3(2)                    'ruc
    wCabe(34) = Date                           'fechafactura
    wCabe(35) = DTGlobal                     'Total Descuentos Globales
    wCabe(36) = DTCliente                    'Total Descuentos Cliente
    wCabe(37) = DTOficina                    'Total Descuentos Oficina
    wCabe(38) = DTItem                       'Total Descuentos Item
    wCabe(39) = DTLinea                      'Total Descuentos Linea
    wCabe(40) = DTPromo                      'Total Descuentos x Promocion
    
    Set asql = VGCNx.Execute("select * from Detallecotizalibre where pedidonumero='" & MBox(1) & "'")
    If asql.RecordCount > 0 Then
       VGCNx.Execute "Delete From Detallecotizalibre where pedidonumero='" & MBox(1) & "'"
       VGCNx.Execute "Delete From cotizalibre where pedidonumero='" & MBox(1) & "'"
    End If
    asql.Close
    nsql = "Insert Into Cotizalibre("
    
    Set asql = Nothing

    wCabe(2) = Trim(MBox(1))                         'nro pedido
    wCabe(3) = Trim(MBox(2))                         'nro factura
    wCabe(4) = Trim(MBox(3))                         'nro boleta
    wCabe(5) = Trim(MBox(4))                         'nro guia
    DoEvents
    '************
    
    Set acmd.ActiveConnection = VGGeneral
    acmd.CommandType = adCmdStoredProc
    acmd.CommandText = "vt_ingresapedido_pro"
    acmd.CommandTimeout = 0
    acmd.Prepared = True
    With acmd
        .Parameters("@base") = VGCNx.DefaultDatabase
        .Parameters("@tabla") = "cotizalibre"
        .Parameters("@tipo") = IIf(dllgeneral.VerificaDatoExistente(VGCNx, "select * from cotizalibre where pedidonumero='" & wCabe(2) & "'") = 0, "1", "2") '"1"
        .Parameters("@puntovta") = wCabe(1)
        .Parameters("@numero") = wCabe(2)
        .Parameters("@factura") = wCabe(3)
        .Parameters("@boleta") = "00"   'wCabe(4)
        .Parameters("@guia") = wCabe(5)
        .Parameters("@dsctoglobal") = wCabe(6)
        .Parameters("@dsctoppago") = wCabe(7)
        .Parameters("@dsctovtaofi") = wCabe(8)
        .Parameters("@moneda") = wCabe(9)
        .Parameters("@tipocambio") = Trim(wCabe(10))
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
        .Parameters("@montodsctoppago") = DTPPago
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
        .Parameters("@observa") = ""
        .Parameters("@tiporefe") = ""
        .Parameters("@nrorefe") = ""
        .Parameters("@nrotransporte") = ""
        .Parameters("@empresa") = VGParametros.empresacodigo
        .Parameters("@TipoContacto") = ""
        .Parameters("@Profesional") = ""
        .Parameters("@hora") = ""
    End With
    acmd.Execute
    
    Set acmd = Nothing
    DoEvents
    '********** DETALLE DE MOVIMIENTOS *****************
    rsdeta.MoveFirst
    regi = 0
    tinafecto = 0
    Do Until rsdeta.EOF
           TCant = rsdeta.Fields(4): TImporte = rsdeta.Fields(5) * rsdeta.Fields(4)
           If IsNull(text1) Or Len(Trim(text1)) = 0 Then
                 dct06 = 0
           Else
               dct06 = TImporte * (CDbl(text1))
           End If
          
           dct02 = 0
           dct02 = (TImporte * (rsdeta.Fields(6) / 100))
           
           dct03 = 0
           dct03 = (TImporte - dct02 - dct06) * (MBox(7) / 100)
            
           dct04 = 0
           dct04 = (TImporte - dct02 - dct03 - dct06) * (MBox(6) / 100)
            
           dct05 = 0
           dct05 = (TImporte - dct02 - dct03 - dct04 - dct06) * (MBox(5) / 100)
           
           Tdscto = dct02 + dct03 + dct04 + dct05 + dct06
            
           TSub = 0
           TSub = (TImporte - (dct02 + dct03 + dct04 + dct05 + dct06))
           Previo = TSub
           If VGParamSistem.tieneigv = "1" Then
              Previo = (TSub / (1 + VGParamSistem.Igv / 100))
           Else
                If modoventa.impuestos = "1" Then
                     Previo = (TImporte - (dct02 + dct03 + dct04 + dct05 + dct06))
                     Previo = (Previo * VGParamSistem.Igv)
                Else
                    If rsdeta.Fields(11) > 0 Then
                         Previo = (TImporte - (dct02 + dct03 + dct04 + dct05 + dct06))
                         Previo = (Previo * rsdeta.Fields(11))
                    Else
                        Previo = TSub '
                        tinafecto = tinafecto + TSub
                   End If
                End If
           End If
        
        nsql = "Detallecotizalibre"
        Set acmd.ActiveConnection = VGGeneral
        acmd.CommandType = adCmdStoredProc
        acmd.CommandTimeout = 0
        acmd.CommandText = "vt_ingresodetallepedido_pro"
        acmd.Prepared = True
        
        With acmd
            .Parameters("@base") = VGCNx.DefaultDatabase
            .Parameters("@tabla") = nsql
            .Parameters("@empresa") = VGParametros.empresacodigo
            .Parameters("@tipo") = "1"
            .Parameters("@item") = rsdeta.Fields(0)
            .Parameters("@numero") = MBox(1)
            .Parameters("@producto") = rsdeta.Fields(1)
            .Parameters("@unidad") = rsdeta.Fields(3)
            .Parameters("@cantidad") = rsdeta.Fields(4)
            .Parameters("@preciopacto") = rsdeta.Fields(7)
            .Parameters("@dsctoxitem") = rsdeta.Fields(6)
            .Parameters("@importebruto") = (rsdeta.Fields(7)) / (1 + VGParamSistem.Igv / 100)
            .Parameters("@porcomision") = rsdeta.Fields(8)
            .Parameters("@mdsctoitem") = Tdscto
            .Parameters("@mdsctoxlinea") = 0
            .Parameters("@mdsctoxprom") = 0     '0
            .Parameters("@mimpor") = rsdeta.Fields(7) - (rsdeta.Fields(7) / (1 + VGParamSistem.Igv / 100)) 'Previo
            .Parameters("@unidadref") = IIf(IsNull(rsdeta.Fields(9)) Or Len(Trim(rsdeta.Fields(9))) = 0, 0, CDbl(rsdeta.Fields(9)))
            .Parameters("@preciolista") = rsdeta.Fields(5)
            .Parameters("@partida") = " "
            .Parameters("@metrica") = " "
            .Parameters("@observacion") = " "
       
        End With
        acmd.Execute
        Set acmd = Nothing
                
        rsdeta.MoveNext
        regi = regi + 1
    Loop
    '*****Actualizamos el Valor de Inafecto**********
    VGCNx.Execute "UPDATE cotizalibre Set Pedidototinafecto=" & tinafecto & _
               " Where pedidonumero='" & MBox(1) & "'"
               
    VGCNx.Execute "Update vt_puntovtadocumento " _
    & " set puntovtadoccorr='" & Right("00000000" & Trim(CStr(MBox(1) + 1)), 8) & "'" _
    & " Where documentocodigo='" & g_tipoped & "' and puntovtacodigo='" & g_ptoventa & "' " _
    & " and puntovtadocserie='" & g_pedserie & "' and empresacodigo='" & VGParametros.empresacodigo & "' " _
    & " and puntovtacodigo='" & VGParametros.puntovta & "'"
  
    MsgBox "Se Grabo Satisfactoriamente el Documento." & Chr(13) & Chr(10) & "Cotizacion => " & MBox(1), vbInformation, MsgTitle
    GrabarData = 1
    'Exit Sub
    
vererror:
   If Err Then
      MsgBox "Comunicarse con Sistemas ...!!" & Err.Description, vbInformation, "Sistemas"
      Exit Function
   End If
End Function


Public Function verificaproducto() As Integer
    On Error Resume Next
    verificaproducto = 0
    If rsdeta.RecordCount > 0 Then
       rsdeta.MoveFirst
       Do Until rsdeta.EOF
           If Escadena(rsdeta.Fields(1)) = MBox2(1) And flag = 0 Then
              verificaproducto = 1
              Exit Do
           End If
           rsdeta.MoveNext
       Loop
    End If
End Function




Public Function CotizaImprimir()
Dim Param(5) As Variant
Dim formulas(4) As Variant
Dim reporte As String
'Dim busca As New dll_apisgen.dll_apis

reporte = "RepCotizacion.rpt"
  
Param(0) = VGParamSistem.BDEmpresa
Param(1) = MBox(1)
Param(2) = VGParametros.empresacodigo
Param(3) = "00"
Param(4) = MBox2(10)

formulas(0) = "@Empresa='" & VGParametros.nomempresa & "'"
formulas(1) = "@Ruc='" & VGParametros.RucEmpresa & "'"
Select Case Left(Combo1.Text, 2)
Case "01":
    formulas(2) = "S/. " & Trim(MBox2(7))
    formulas(3) = "S/. " & Trim(MBox2(9))
Case "02":
    formulas(2) = "US$/. " & Trim(MBox2(7))
    formulas(3) = "US$/. " & Trim(MBox2(9))
Case "03":
    formulas(2) = "e/. " & Trim(MBox2(7))
    formulas(3) = "e/. " & Trim(MBox2(9))
End Select

Call ImpresionRptProc(reporte, formulas, Param, , "Cotizacion")
   
End Function




Public Sub TraerProducto()
  Dim rabusca As New ADODB.Recordset
  Dim mone As String
  Dim nprecio As Double
  Dim nsql As String

   On Error Resume Next
   
    If Combo2.ListCount > 0 Then
       nsql = "select *,stskdis from [" & VGCNx.DefaultDatabase & "].dbo.maeart " & _
              " inner join [" & _
              VGCNx.DefaultDatabase & "].dbo.stkart " & _
              " ON acodigo=stcodigo " & _
              " where acodigo='" & MBox2(1) & "' and stalma='" & Ctr_Ayuda3.xclave & "'"
    Else
       nsql = "select *,stskdis from [" & VGCNx.DefaultDatabase & "].dbo.maeart " & _
              " inner join [" & _
              VGCNx.DefaultDatabase & "].dbo.stkart " & _
              " ON acodigo=stcodigo " & _
              " where acodigo='" & MBox2(1) & "' and stalma='" & Ctr_Ayuda3.xclave & "'"
    End If
    Set rabusca = VGCNx.Execute(nsql)
    If rabusca.RecordCount > 0 Then
      Label2 = Escadena(rabusca!adescri)
      MBox2(2) = Escadena(rabusca!unidadcodigo)
      If rabusca!acodmon = "01" Then
        mone = g_tiposol
      ElseIf rabusca!acodmon = "02" Then
        mone = g_tipodolar
      Else
         mone = rabusca!acodmon
      End If
      If mone <> dllgeneral.ComboDato(Combo1.Text) Then
         If dllgeneral.ComboDato(Combo1.Text) = g_tiposol Then
            nprecio = TraePrecio(Combo2.Text, MBox2(1), VGCNx, Trim(Ctr_Ayuda3.xclave))
            If nprecio > 0 Then
               MBox2(3) = numero(nprecio * CDbl(MBox(8)))
               If modoventa.unidadmedida = "R" And modoventa.usafactor = "1" Then
                  MBox2(0) = numero(0)     'rabusca!unidadfactorconv)
                  MBox2(13) = numero(0)    'rabusca!unidadfactorconv)
               ElseIf modoventa.unidadmedida = "V" And modoventa.usafactor = "1" Then
                  MBox2(13) = numero(0)    'rabusca!unidadfactorconv)
               Else
                  MBox2(13) = 1
               End If
               MBox2(14) = numero(0)    'rabusca!productoporcimpto)
            Else
               MBox2(3) = numero(TraePrecio(Combo2.Text, MBox2(1), VGCNx, Trim(Ctr_Ayuda3.xclave))) 'rabusca!productoprecvta)
               If modoventa.unidadmedida = "R" And modoventa.usafactor = "1" Then
                  MBox2(0) = numero(0)   'rabusca!unidadfactorconv)
                  MBox2(13) = numero(0)  'rabusca!unidadfactorconv)
               ElseIf modoventa.unidadmedida = "V" And modoventa.usafactor = "1" Then
                  MBox2(13) = numero(0)  'rabusca!unidadfactorconv)
               Else
                  MBox2(13) = 1
               End If
               MBox2(14) = numero(0)    'rabusca!productoporcimpto)
            End If
         ElseIf dllgeneral.ComboDato(Combo1.Text) = g_tipodolar Then
            nprecio = TraePrecio(Combo2.Text, MBox2(1), VGCNx, Trim(Ctr_Ayuda3.xclave))
            If nprecio > 0 Then
               MBox2(3) = numero(nprecio / CDbl(MBox(8)))
               If modoventa.unidadmedida = "R" And modoventa.usafactor = "1" Then
                  MBox2(0) = numero(0)    'rabusca!unidadfactorconv)
                  MBox2(13) = numero(0)   'rabusca!unidadfactorconv)
               ElseIf modoventa.unidadmedida = "V" And modoventa.usafactor = "1" Then
                  MBox2(13) = numero(0)    'rabusca!unidadfactorconv)
               Else
                  MBox2(13) = 1
               End If
               MBox2(14) = numero(0)     'rabusca!productoporcimpto)
            Else
               MBox2(3) = numero(TraePrecio(Combo2.Text, MBox2(1), VGCNx, Trim(Ctr_Ayuda3.xclave))) 'rabusca!productoprecvta)
               If modoventa.unidadmedida = "R" And modoventa.usafactor = "1" Then
                  MBox2(0) = numero(0)    'rabusca!unidadfactorconv)
                  MBox2(13) = numero(0)   'rabusca!unidadfactorconv)
               ElseIf modoventa.unidadmedida = "V" And modoventa.usafactor = "1" Then
                  MBox2(13) = numero(0)   'rabusca!unidadfactorconv)
               Else
                  MBox2(13) = 1
               End If
               MBox2(14) = numero(0)      'rabusca!productoporcimpto)
            End If
         End If
      Else
         MBox2(3) = numero(TraePrecio(Combo2.Text, MBox2(1), VGCNx, Trim(Ctr_Ayuda3.xclave)))  'rabusca!productoprecvta)
         If modoventa.unidadmedida = "R" And modoventa.usafactor = "1" Then
            MBox2(0) = numero(0)    'rabusca!unidadfactorconv)
            MBox2(13) = numero(0)   'rabusca!unidadfactorconv)
         ElseIf modoventa.unidadmedida = "V" And modoventa.usafactor = "1" Then
            MBox2(13) = numero(0)   'rabusca!unidadfactorconv)
         Else
            MBox2(13) = 1
         End If
         MBox2(14) = numero(0)    'rabusca!productoporcimpto)
      End If
    Else
      Label2 = "":    MBox2(2) = ""
    End If
    MBox2(4) = numero(0)
    MBox2(5) = numero(0)
    rabusca.Close
    Set rabusca = Nothing
End Sub

Private Sub TxtCliente_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub


Private Sub TxtNro_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub


