VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "TODG7.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "CONTROLAYUDA.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form FrmPedido 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pedido"
   ClientHeight    =   7920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7920
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.StatusBar Panel 
      Align           =   2  'Align Bottom
      Height          =   312
      Left            =   0
      TabIndex        =   1
      Top             =   7608
      Visible         =   0   'False
      Width           =   11904
      _ExtentX        =   21008
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7650
      Left            =   0
      TabIndex        =   0
      Top             =   30
      Width           =   11925
      _ExtentX        =   21034
      _ExtentY        =   13494
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      BackColor       =   12632256
      TabCaption(0)   =   "Consulta"
      TabPicture(0)   =   "XXFrmPedidoDETALLADO.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Fr1(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "frmbotones"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Mantenimiento"
      TabPicture(1)   =   "XXFrmPedidoDETALLADO.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).Control(1)=   "Fr4"
      Tab(1).Control(2)=   "SSTab2"
      Tab(1).Control(3)=   "Frame4"
      Tab(1).Control(4)=   "Fr2(2)"
      Tab(1).Control(5)=   "TDBGrid1"
      Tab(1).Control(6)=   "Frame6"
      Tab(1).Control(7)=   "Fr2(1)"
      Tab(1).ControlCount=   8
      TabCaption(2)   =   "Ingreso Masivo"
      TabPicture(2)   =   "XXFrmPedidoDETALLADO.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label1(26)"
      Tab(2).Control(1)=   "Label3(6)"
      Tab(2).Control(2)=   "Label4(2)"
      Tab(2).Control(3)=   "TDBGrid3"
      Tab(2).Control(4)=   "Text10"
      Tab(2).Control(5)=   "Text7"
      Tab(2).Control(6)=   "Text4(3)"
      Tab(2).Control(7)=   "Cmdgrabamasivo"
      Tab(2).Control(8)=   "Cmdsalirmasivo"
      Tab(2).ControlCount=   9
      Begin VB.CommandButton Cmdsalirmasivo 
         Caption         =   "Cancelar"
         Height          =   540
         Left            =   -64695
         TabIndex        =   165
         Top             =   6525
         Width           =   972
      End
      Begin VB.CommandButton Cmdgrabamasivo 
         Caption         =   "Grabar"
         Height          =   540
         Left            =   -66045
         TabIndex        =   164
         Top             =   6570
         Width           =   972
      End
      Begin VB.TextBox Text4 
         Height          =   396
         Index           =   3
         Left            =   -69096
         TabIndex        =   163
         Text            =   "Text4"
         Top             =   6645
         Visible         =   0   'False
         Width           =   972
      End
      Begin VB.TextBox Text7 
         Height          =   396
         Left            =   -71064
         TabIndex        =   162
         Text            =   "Text3"
         Top             =   6705
         Visible         =   0   'False
         Width           =   972
      End
      Begin VB.TextBox Text10 
         Height          =   396
         Left            =   -74712
         TabIndex        =   161
         Text            =   "Text1"
         Top             =   6705
         Width           =   972
      End
      Begin VB.Frame frmbotones 
         Height          =   930
         Left            =   4440
         TabIndex        =   148
         Top             =   6768
         Width           =   4500
         Begin VB.CommandButton cmdBotones 
            Caption         =   "&Salir"
            Height          =   690
            Index           =   4
            Left            =   3375
            Picture         =   "XXFrmPedidoDETALLADO.frx":0054
            Style           =   1  'Graphical
            TabIndex        =   152
            Top             =   180
            Width           =   870
         End
         Begin VB.CommandButton cmdBotones 
            Caption         =   "&Eliminar"
            Height          =   690
            Index           =   2
            Left            =   2340
            Picture         =   "XXFrmPedidoDETALLADO.frx":0496
            Style           =   1  'Graphical
            TabIndex        =   151
            Top             =   180
            Width           =   825
         End
         Begin VB.CommandButton cmdBotones 
            Caption         =   "E&ditar"
            Height          =   690
            Index           =   1
            Left            =   1260
            Picture         =   "XXFrmPedidoDETALLADO.frx":08D8
            Style           =   1  'Graphical
            TabIndex        =   150
            Top             =   180
            Width           =   870
         End
         Begin VB.CommandButton cmdBotones 
            Caption         =   "&Nuevo"
            Height          =   690
            Index           =   0
            Left            =   180
            Picture         =   "XXFrmPedidoDETALLADO.frx":0D1A
            Style           =   1  'Graphical
            TabIndex        =   149
            Top             =   180
            Width           =   870
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
         Height          =   735
         Left            =   -74880
         TabIndex        =   135
         Top             =   6948
         Width           =   2835
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
            TabIndex        =   139
            Top             =   450
            Width           =   675
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
            TabIndex        =   138
            Top             =   150
            Width           =   495
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   150
            Picture         =   "XXFrmPedidoDETALLADO.frx":115C
            Top             =   210
            Width           =   480
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
            TabIndex        =   137
            Top             =   450
            Width           =   1125
         End
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
            TabIndex        =   136
            Top             =   150
            Width           =   1095
         End
      End
      Begin VB.Frame Fr4 
         BorderStyle     =   0  'None
         Height          =   1755
         Left            =   -70950
         TabIndex        =   123
         Top             =   4488
         Visible         =   0   'False
         Width           =   3825
         Begin VB.Frame Frame2 
            BackColor       =   &H00800000&
            BorderStyle     =   0  'None
            Height          =   1365
            Left            =   300
            TabIndex        =   126
            Top             =   240
            Width           =   1785
            Begin VB.OptionButton cOpc2 
               BackColor       =   &H00800000&
               Caption         =   "BO"
               ForeColor       =   &H00C0FFC0&
               Height          =   285
               Index           =   2
               Left            =   360
               TabIndex        =   129
               Top             =   960
               Width           =   975
            End
            Begin VB.OptionButton cOpc2 
               BackColor       =   &H00800000&
               Caption         =   "Boleta"
               ForeColor       =   &H00C0FFC0&
               Height          =   285
               Index           =   1
               Left            =   330
               TabIndex        =   128
               Top             =   570
               Width           =   945
            End
            Begin VB.OptionButton cOpc2 
               BackColor       =   &H00800000&
               Caption         =   "Factura"
               ForeColor       =   &H00C0FFC0&
               Height          =   285
               Index           =   0
               Left            =   300
               TabIndex        =   127
               Top             =   210
               Value           =   -1  'True
               Width           =   945
            End
         End
         Begin VB.CommandButton cSeleccion 
            BackColor       =   &H0000C0C0&
            Caption         =   "Canc&ela"
            Height          =   435
            Index           =   1
            Left            =   2190
            Style           =   1  'Graphical
            TabIndex        =   125
            Top             =   1020
            Width           =   1245
         End
         Begin VB.CommandButton cSeleccion 
            BackColor       =   &H0000C0C0&
            Caption         =   "Ace&pta"
            Height          =   435
            Index           =   0
            Left            =   2190
            Style           =   1  'Graphical
            TabIndex        =   124
            Top             =   480
            Width           =   1245
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00800000&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H80000006&
            Height          =   1395
            Index           =   5
            Left            =   150
            Shape           =   4  'Rounded Rectangle
            Top             =   240
            Width           =   3495
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00C0FFC0&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H80000006&
            Height          =   1635
            Index           =   6
            Left            =   0
            Shape           =   4  'Rounded Rectangle
            Top             =   120
            Width           =   3795
         End
      End
      Begin VB.Frame Fr1 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1815
         Index           =   1
         Left            =   4500
         TabIndex        =   99
         Top             =   3948
         Visible         =   0   'False
         Width           =   3855
         Begin VB.CommandButton cBoton 
            BackColor       =   &H0000C0C0&
            Caption         =   "&Cancela"
            Height          =   435
            Index           =   1
            Left            =   2070
            Style           =   1  'Graphical
            TabIndex        =   103
            Top             =   960
            Width           =   1305
         End
         Begin VB.CommandButton cBoton 
            BackColor       =   &H0000C0C0&
            Caption         =   "&Acepta"
            Height          =   435
            Index           =   0
            Left            =   570
            Style           =   1  'Graphical
            TabIndex        =   102
            Top             =   960
            Width           =   1305
         End
         Begin VB.OptionButton cOpc 
            BackColor       =   &H00800000&
            Caption         =   "&FACTURACION"
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
            Left            =   1980
            MaskColor       =   &H00FFFFFF&
            TabIndex        =   101
            Top             =   480
            Width           =   1665
         End
         Begin VB.OptionButton cOpc 
            BackColor       =   &H00800000&
            Caption         =   "&PEDIDOS"
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
            Left            =   390
            TabIndex        =   100
            Top             =   480
            Value           =   -1  'True
            Width           =   1485
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00800000&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H80000006&
            Height          =   1485
            Index           =   0
            Left            =   120
            Shape           =   4  'Rounded Rectangle
            Top             =   120
            Width           =   3615
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00FFFFFF&
            BorderStyle     =   6  'Inside Solid
            FillColor       =   &H00C0FFC0&
            FillStyle       =   0  'Solid
            Height          =   1695
            Index           =   1
            Left            =   30
            Shape           =   4  'Rounded Rectangle
            Top             =   0
            Width           =   3795
         End
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   2508
         Left            =   -74940
         TabIndex        =   26
         Top             =   684
         Width           =   11808
         _ExtentX        =   20823
         _ExtentY        =   4419
         _Version        =   393216
         Style           =   1
         TabHeight       =   520
         TabCaption(0)   =   "Datos Generales"
         TabPicture(0)   =   "XXFrmPedidoDETALLADO.frx":159E
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Fr1(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Datos Detalle"
         TabPicture(1)   =   "XXFrmPedidoDETALLADO.frx":15BA
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Chkmasivo"
         Tab(1).Control(1)=   "TClie"
         Tab(1).Control(2)=   "Fr2(0)"
         Tab(1).ControlCount=   3
         TabCaption(2)   =   "Datos Complementarios"
         TabPicture(2)   =   "XXFrmPedidoDETALLADO.frx":15D6
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Fr3(0)"
         Tab(2).ControlCount=   1
         Begin VB.CheckBox Chkmasivo 
            Caption         =   "Ing.Masivo"
            Height          =   192
            Left            =   -66825
            TabIndex        =   167
            Top             =   2256
            Value           =   1  'Checked
            Width           =   1290
         End
         Begin VB.CheckBox TClie 
            Caption         =   "Cliente Eventual"
            Height          =   195
            Left            =   -65040
            TabIndex        =   122
            Top             =   2250
            Width           =   1515
         End
         Begin VB.Frame Fr1 
            Height          =   1815
            Index           =   0
            Left            =   150
            TabIndex        =   68
            Top             =   360
            Width           =   11565
            Begin VB.ComboBox Combo1 
               Height          =   288
               Left            =   3456
               Style           =   2  'Dropdown List
               TabIndex        =   77
               Top             =   750
               Width           =   1308
            End
            Begin VB.ComboBox Combo2 
               Height          =   288
               Left            =   10296
               Style           =   2  'Dropdown List
               TabIndex        =   79
               Top             =   804
               Width           =   1065
            End
            Begin MSMask.MaskEdBox MBox 
               Height          =   255
               Index           =   0
               Left            =   1410
               TabIndex        =   69
               Top             =   420
               Width           =   375
               _ExtentX        =   661
               _ExtentY        =   450
               _Version        =   393216
               MaxLength       =   2
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MBox 
               Height          =   255
               Index           =   1
               Left            =   2730
               TabIndex        =   70
               Top             =   390
               Width           =   1125
               _ExtentX        =   1984
               _ExtentY        =   450
               _Version        =   393216
               MaxLength       =   11
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MBox 
               Height          =   255
               Index           =   2
               Left            =   4935
               TabIndex        =   71
               Top             =   420
               Width           =   1125
               _ExtentX        =   1984
               _ExtentY        =   450
               _Version        =   393216
               ClipMode        =   1
               Enabled         =   0   'False
               MaxLength       =   11
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MBox 
               Height          =   255
               Index           =   3
               Left            =   6930
               TabIndex        =   72
               Top             =   435
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   450
               _Version        =   393216
               ClipMode        =   1
               Enabled         =   0   'False
               MaxLength       =   11
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MBox 
               Height          =   255
               Index           =   4
               Left            =   8955
               TabIndex        =   73
               Top             =   390
               Width           =   1275
               _ExtentX        =   2249
               _ExtentY        =   450
               _Version        =   393216
               ClipMode        =   1
               Enabled         =   0   'False
               MaxLength       =   11
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MBox 
               Height          =   255
               Index           =   5
               Left            =   1410
               TabIndex        =   74
               Top             =   810
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   450
               _Version        =   393216
               ClipMode        =   1
               MaxLength       =   9
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MBox 
               Height          =   252
               Index           =   6
               Left            =   1392
               TabIndex        =   75
               Top             =   1488
               Visible         =   0   'False
               Width           =   972
               _ExtentX        =   1720
               _ExtentY        =   450
               _Version        =   393216
               ClipMode        =   1
               MaxLength       =   9
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MBox 
               Height          =   252
               Index           =   7
               Left            =   3732
               TabIndex        =   76
               Top             =   1488
               Visible         =   0   'False
               Width           =   972
               _ExtentX        =   1720
               _ExtentY        =   450
               _Version        =   393216
               ClipMode        =   1
               MaxLength       =   9
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MBox 
               Height          =   252
               Index           =   8
               Left            =   6852
               TabIndex        =   78
               Top             =   780
               Width           =   972
               _ExtentX        =   1720
               _ExtentY        =   450
               _Version        =   393216
               ClipMode        =   1
               MaxLength       =   9
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MBox 
               Height          =   255
               Index           =   9
               Left            =   3510
               TabIndex        =   80
               Top             =   1200
               Width           =   7965
               _ExtentX        =   14049
               _ExtentY        =   450
               _Version        =   393216
               MaxLength       =   45
               PromptChar      =   "_"
            End
            Begin VB.Label Label1 
               Caption         =   "Punto Venta"
               Height          =   255
               Index           =   0
               Left            =   240
               TabIndex        =   92
               Top             =   420
               Width           =   975
            End
            Begin VB.Label Label1 
               Caption         =   "No .Factura"
               Height          =   255
               Index           =   1
               Left            =   3975
               TabIndex        =   91
               Top             =   420
               Width           =   1215
            End
            Begin VB.Label Label1 
               Caption         =   "Dcto. Genral."
               Height          =   255
               Index           =   2
               Left            =   240
               TabIndex        =   90
               Top             =   810
               Width           =   1245
            End
            Begin VB.Label Label1 
               Caption         =   "Tipo de Cambio"
               Height          =   252
               Index           =   3
               Left            =   5592
               TabIndex        =   89
               Top             =   816
               Width           =   1212
            End
            Begin VB.Label Label1 
               Caption         =   "No .Boleta"
               Height          =   255
               Index           =   4
               Left            =   6090
               TabIndex        =   88
               Top             =   450
               Width           =   975
            End
            Begin VB.Label Label1 
               Caption         =   "Dcto. Promoc."
               Height          =   252
               Index           =   5
               Left            =   288
               TabIndex        =   87
               Top             =   1488
               Visible         =   0   'False
               Width           =   1032
            End
            Begin VB.Label Label1 
               Caption         =   "Moneda"
               Height          =   252
               Index           =   6
               Left            =   2676
               TabIndex        =   86
               Top             =   840
               Width           =   1032
            End
            Begin VB.Label Label1 
               Caption         =   "No. Pedido"
               Height          =   255
               Index           =   7
               Left            =   1890
               TabIndex        =   85
               Top             =   420
               Width           =   855
            End
            Begin VB.Label Label1 
               Caption         =   "No. Guia"
               Height          =   255
               Index           =   8
               Left            =   8265
               TabIndex        =   84
               Top             =   420
               Width           =   855
            End
            Begin VB.Label Label1 
               Caption         =   "Dcto. Especial"
               Height          =   252
               Index           =   9
               Left            =   2592
               TabIndex        =   83
               Top             =   1488
               Visible         =   0   'False
               Width           =   1128
            End
            Begin VB.Label Label1 
               Caption         =   "Lista Precios"
               Height          =   252
               Index           =   10
               Left            =   9120
               TabIndex        =   82
               Top             =   864
               Width           =   1212
            End
            Begin VB.Label Label1 
               Caption         =   "Mensajes"
               Height          =   252
               Index           =   11
               Left            =   2640
               TabIndex        =   81
               Top             =   1200
               Width           =   1212
            End
         End
         Begin VB.Frame Fr2 
            Height          =   1875
            Index           =   0
            Left            =   -74952
            TabIndex        =   27
            Top             =   330
            Width           =   11685
            Begin VB.TextBox Text5 
               Enabled         =   0   'False
               Height          =   288
               Left            =   10032
               TabIndex        =   157
               Text            =   "Text1"
               Top             =   528
               Width           =   1500
            End
            Begin VB.TextBox Text2 
               Enabled         =   0   'False
               Height          =   285
               Left            =   10560
               TabIndex        =   147
               Top             =   1488
               Width           =   1005
            End
            Begin VB.TextBox text1 
               Enabled         =   0   'False
               Height          =   285
               Left            =   10530
               TabIndex        =   146
               Top             =   1470
               Width           =   1005
            End
            Begin VB.CommandButton cAyuda 
               Caption         =   "..."
               Height          =   255
               Index           =   0
               Left            =   9108
               TabIndex        =   43
               Top             =   1200
               Width           =   285
            End
            Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuda3 
               Height          =   312
               Left            =   6588
               TabIndex        =   34
               Top             =   840
               Width           =   2628
               _ExtentX        =   4630
               _ExtentY        =   556
               XcodMaxLongitud =   2
               xcodwith        =   100
               NomTabla        =   "vt_almacen"
               TituloAyuda     =   "Ayuda de Almacenes"
               ListaCampos     =   "almacencodigo(1),almacendescripcion(1)"
               XcodCampo       =   "almacencodigo"
               XListCampo      =   "almacendescripcion"
               ListaCamposDescrip=   "Codigo,Descripcion"
               ListaCamposText =   "almacencodigo,almacendescripcion"
            End
            Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuda2 
               Height          =   312
               Left            =   1896
               TabIndex        =   32
               Top             =   828
               Width           =   3792
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
            Begin VB.ComboBox Combo3 
               Height          =   315
               Left            =   1890
               Style           =   2  'Dropdown List
               TabIndex        =   28
               Top             =   180
               Width           =   2265
            End
            Begin VB.ComboBox Combo4 
               Height          =   288
               Left            =   9264
               Style           =   2  'Dropdown List
               TabIndex        =   30
               Top             =   180
               Width           =   2445
            End
            Begin VB.ComboBox Combo5 
               Height          =   288
               Left            =   10596
               Style           =   2  'Dropdown List
               TabIndex        =   39
               Top             =   840
               Width           =   735
            End
            Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuda1 
               Height          =   312
               Left            =   1896
               TabIndex        =   31
               Top             =   516
               Width           =   6576
               _ExtentX        =   11589
               _ExtentY        =   556
               XcodMaxLongitud =   11
               xcodwith        =   800
               NomTabla        =   "vt_Cliente"
               TituloAyuda     =   "Ayuda de Clientes"
               ListaCampos     =   $"XXFrmPedidoDETALLADO.frx":15F2
               XcodCampo       =   "clientecodigo"
               XListCampo      =   "clienterazonsocial"
               ListaCamposDescrip=   "Codigo,Descripcion,Ruc,Direccion,Distrito,LimiteCred,Saldo,T,P,M,D"
               ListaCamposText =   $"XXFrmPedidoDETALLADO.frx":16D8
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
               Height          =   255
               Index           =   10
               Left            =   6120
               TabIndex        =   29
               Top             =   240
               Width           =   1155
               _ExtentX        =   2037
               _ExtentY        =   450
               _Version        =   393216
               ClipMode        =   1
               AllowPrompt     =   -1  'True
               MaxLength       =   10
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MBox 
               Height          =   252
               Index           =   13
               Left            =   8868
               TabIndex        =   33
               Top             =   1500
               Visible         =   0   'False
               Width           =   612
               _ExtentX        =   1085
               _ExtentY        =   450
               _Version        =   393216
               ClipMode        =   1
               MaxLength       =   10
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MBox 
               Height          =   252
               Index           =   15
               Left            =   1896
               TabIndex        =   35
               Top             =   1512
               Visible         =   0   'False
               Width           =   1188
               _ExtentX        =   2117
               _ExtentY        =   450
               _Version        =   393216
               ClipMode        =   1
               MaxLength       =   10
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MBox 
               Height          =   252
               Index           =   16
               Left            =   4284
               TabIndex        =   36
               Top             =   1512
               Visible         =   0   'False
               Width           =   1188
               _ExtentX        =   2090
               _ExtentY        =   450
               _Version        =   393216
               ClipMode        =   1
               MaxLength       =   10
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MBox 
               Height          =   252
               Index           =   17
               Left            =   6876
               TabIndex        =   37
               Top             =   1512
               Visible         =   0   'False
               Width           =   1032
               _ExtentX        =   1826
               _ExtentY        =   450
               _Version        =   393216
               ClipMode        =   1
               MaxLength       =   10
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MBox 
               Height          =   255
               Index           =   18
               Left            =   11010
               TabIndex        =   41
               Top             =   1170
               Width           =   525
               _ExtentX        =   926
               _ExtentY        =   450
               _Version        =   393216
               ClipMode        =   1
               MaxLength       =   3
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MBox 
               Height          =   252
               Index           =   19
               Left            =   1896
               TabIndex        =   42
               Top             =   1200
               Width           =   7116
               _ExtentX        =   12568
               _ExtentY        =   450
               _Version        =   393216
               PromptChar      =   "_"
            End
            Begin VB.Label Label1 
               Caption         =   "RUC"
               Height          =   252
               Index           =   25
               Left            =   8928
               TabIndex        =   156
               Top             =   528
               Width           =   972
            End
            Begin VB.Label Label6 
               Caption         =   "Dscto Cliente"
               Height          =   255
               Left            =   9525
               TabIndex        =   145
               Top             =   1545
               Width           =   975
            End
            Begin VB.Label Label1 
               Caption         =   "Punto de Llegada"
               Height          =   252
               Index           =   24
               Left            =   216
               TabIndex        =   140
               Top             =   1164
               Width           =   1332
            End
            Begin VB.Label Label1 
               Caption         =   "Dias Pago"
               Height          =   255
               Index           =   18
               Left            =   10170
               TabIndex        =   67
               Top             =   1200
               Width           =   765
            End
            Begin VB.Label Label1 
               Caption         =   "Modo de la Venta"
               Height          =   255
               Index           =   12
               Left            =   240
               TabIndex        =   66
               Top             =   240
               Width           =   1575
            End
            Begin VB.Label Label1 
               Caption         =   "Fecha de Atencion"
               Height          =   225
               Index           =   13
               Left            =   4590
               TabIndex        =   65
               Top             =   240
               Width           =   1365
            End
            Begin VB.Label Label1 
               Caption         =   "Forma de Pago"
               Height          =   252
               Index           =   14
               Left            =   7872
               TabIndex        =   64
               Top             =   240
               Width           =   1140
            End
            Begin VB.Label Label1 
               Caption         =   "Codigo del Cliente"
               Height          =   255
               Index           =   15
               Left            =   240
               TabIndex        =   63
               Top             =   570
               Width           =   1575
            End
            Begin VB.Label Label1 
               Caption         =   "Codigo del Vendedor"
               Height          =   252
               Index           =   16
               Left            =   240
               TabIndex        =   62
               Top             =   900
               Width           =   1572
            End
            Begin VB.Label Label1 
               Caption         =   "Almacen"
               Height          =   252
               Index           =   17
               Left            =   5784
               TabIndex        =   61
               Top             =   900
               Width           =   792
            End
            Begin VB.Label Label1 
               Caption         =   "Otros Gastos"
               Height          =   252
               Index           =   19
               Left            =   240
               TabIndex        =   60
               Top             =   1536
               Visible         =   0   'False
               Width           =   1572
            End
            Begin VB.Label Label1 
               Caption         =   "Nota de Pedido"
               Height          =   252
               Index           =   20
               Left            =   3096
               TabIndex        =   58
               Top             =   1512
               Visible         =   0   'False
               Width           =   1140
            End
            Begin VB.Label Label1 
               Caption         =   "Orden de Compra"
               Height          =   252
               Index           =   21
               Left            =   5520
               TabIndex        =   57
               Top             =   1512
               Visible         =   0   'False
               Width           =   1392
            End
            Begin VB.Label Label1 
               Caption         =   "Autorizacion"
               Height          =   252
               Index           =   22
               Left            =   9612
               TabIndex        =   40
               Top             =   864
               Width           =   948
            End
            Begin VB.Label Label1 
               Caption         =   "% Comision"
               Height          =   252
               Index           =   23
               Left            =   7968
               TabIndex        =   38
               Top             =   1500
               Visible         =   0   'False
               Width           =   852
            End
         End
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
            TabIndex        =   93
            Top             =   450
            Width           =   11565
            Begin VB.ComboBox Combo8 
               Height          =   315
               Left            =   1290
               Style           =   2  'Dropdown List
               TabIndex        =   117
               Top             =   1290
               Width           =   1185
            End
            Begin VB.ComboBox Combo7 
               Height          =   315
               Left            =   9540
               Style           =   2  'Dropdown List
               TabIndex        =   115
               Top             =   930
               Width           =   1410
            End
            Begin VB.ComboBox Combo6 
               Height          =   315
               Left            =   7320
               Style           =   2  'Dropdown List
               TabIndex        =   114
               Top             =   930
               Width           =   1125
            End
            Begin MSMask.MaskEdBox MBox3 
               Height          =   315
               Index           =   0
               Left            =   1290
               TabIndex        =   104
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
               TabIndex        =   105
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
               TabIndex        =   106
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
               TabIndex        =   107
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
               TabIndex        =   108
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
               Caption         =   "Multidireccion"
               ForeColor       =   &H80000006&
               Height          =   255
               Index           =   8
               Left            =   240
               TabIndex        =   116
               Top             =   1380
               Width           =   1005
            End
            Begin VB.Label Dclie 
               BackStyle       =   0  'Transparent
               Caption         =   "Pais"
               ForeColor       =   &H80000006&
               Height          =   255
               Index           =   6
               Left            =   9030
               TabIndex        =   113
               Top             =   990
               Width           =   465
            End
            Begin VB.Label Dclie 
               BackStyle       =   0  'Transparent
               Caption         =   "Tipo Persona"
               ForeColor       =   &H80000006&
               Height          =   255
               Index           =   3
               Left            =   6120
               TabIndex        =   112
               Top             =   960
               Width           =   975
            End
            Begin VB.Label Dclie 
               BackStyle       =   0  'Transparent
               Caption         =   "Ruc"
               ForeColor       =   &H80000006&
               Height          =   255
               Index           =   7
               Left            =   9420
               TabIndex        =   111
               Top             =   270
               Width           =   675
            End
            Begin VB.Label lcred 
               Alignment       =   1  'Right Justify
               ForeColor       =   &H8000000E&
               Height          =   285
               Index           =   1
               Left            =   9870
               TabIndex        =   110
               Top             =   1320
               Width           =   1605
            End
            Begin VB.Label lcred 
               Alignment       =   1  'Right Justify
               ForeColor       =   &H8000000C&
               Height          =   285
               Index           =   0
               Left            =   6780
               TabIndex        =   109
               Top             =   1350
               Width           =   1575
            End
            Begin VB.Label Dclie 
               BackStyle       =   0  'Transparent
               Caption         =   "Cliente"
               ForeColor       =   &H80000006&
               Height          =   255
               Index           =   0
               Left            =   240
               TabIndex        =   98
               Top             =   240
               Width           =   1815
            End
            Begin VB.Label Dclie 
               BackStyle       =   0  'Transparent
               Caption         =   "Direccion"
               ForeColor       =   &H80000006&
               Height          =   255
               Index           =   1
               Left            =   240
               TabIndex        =   97
               Top             =   600
               Width           =   1815
            End
            Begin VB.Label Dclie 
               BackStyle       =   0  'Transparent
               Caption         =   "Distrito"
               ForeColor       =   &H80000006&
               Height          =   255
               Index           =   2
               Left            =   240
               TabIndex        =   96
               Top             =   990
               Width           =   1815
            End
            Begin VB.Label Dclie 
               BackStyle       =   0  'Transparent
               Caption         =   "Saldo US$"
               ForeColor       =   &H80000006&
               Height          =   255
               Index           =   4
               Left            =   5790
               TabIndex        =   95
               Top             =   1380
               Width           =   1335
            End
            Begin VB.Label Dclie 
               BackStyle       =   0  'Transparent
               Caption         =   "Limite Cred US$"
               ForeColor       =   &H80000006&
               Height          =   255
               Index           =   5
               Left            =   8520
               TabIndex        =   94
               Top             =   1380
               Width           =   1815
            End
         End
      End
      Begin VB.Frame Frame4 
         Height          =   930
         Left            =   -68970
         TabIndex        =   25
         Top             =   6948
         Width           =   1980
         Begin VB.CommandButton cmdBotones 
            Caption         =   "&Cancelar"
            Height          =   690
            Index           =   12
            Left            =   1050
            Picture         =   "XXFrmPedidoDETALLADO.frx":179D
            Style           =   1  'Graphical
            TabIndex        =   56
            Top             =   180
            Width           =   855
         End
         Begin VB.CommandButton cmdBotones 
            Caption         =   "&Acepta"
            Height          =   690
            Index           =   11
            Left            =   90
            Picture         =   "XXFrmPedidoDETALLADO.frx":1BDF
            Style           =   1  'Graphical
            TabIndex        =   55
            Top             =   180
            Width           =   870
         End
      End
      Begin VB.Frame Fr2 
         BackColor       =   &H00800000&
         BorderStyle     =   0  'None
         Height          =   675
         Index           =   2
         Left            =   -74790
         TabIndex        =   14
         Top             =   6258
         Width           =   11535
         Begin MSMask.MaskEdBox MBox2 
            Height          =   375
            Index           =   6
            Left            =   300
            TabIndex        =   15
            Top             =   120
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
            TabIndex        =   16
            Top             =   120
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
            TabIndex        =   17
            Top             =   120
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
            TabIndex        =   18
            Top             =   120
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
            TabIndex        =   19
            Top             =   120
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
            ForeColor       =   &H0080FF80&
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   24
            Top             =   480
            Width           =   1335
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
            ForeColor       =   &H0080FF80&
            Height          =   255
            Index           =   1
            Left            =   2760
            TabIndex        =   23
            Top             =   480
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
            ForeColor       =   &H0080FF80&
            Height          =   255
            Index           =   2
            Left            =   5160
            TabIndex        =   22
            Top             =   480
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
            ForeColor       =   &H0080FF80&
            Height          =   255
            Index           =   3
            Left            =   7680
            TabIndex        =   21
            Top             =   480
            Width           =   1215
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
            ForeColor       =   &H0080FF80&
            Height          =   255
            Index           =   4
            Left            =   9840
            TabIndex        =   20
            Top             =   480
            Width           =   1215
         End
         Begin VB.Line Line2 
            BorderColor     =   &H80000005&
            Index           =   0
            X1              =   2175
            X2              =   2175
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
            Index           =   2
            X1              =   6960
            X2              =   6960
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
            BorderColor     =   &H00808080&
            Index           =   4
            X1              =   2160
            X2              =   2160
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
            Index           =   6
            X1              =   6940
            X2              =   6940
            Y1              =   120
            Y2              =   1215
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00808080&
            Index           =   7
            X1              =   9340
            X2              =   9340
            Y1              =   120
            Y2              =   1215
         End
      End
      Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
         Height          =   1635
         Left            =   -74940
         TabIndex        =   2
         Top             =   4593
         Width           =   11775
         _ExtentX        =   20770
         _ExtentY        =   2884
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
      Begin VB.Frame Frame1 
         Height          =   5775
         Left            =   210
         TabIndex        =   118
         Top             =   948
         Width           =   11535
         Begin VB.Frame Fr5 
            BorderStyle     =   0  'None
            Caption         =   "TIPO TRANSACCION"
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
            Height          =   1725
            Left            =   4110
            TabIndex        =   130
            Top             =   2250
            Visible         =   0   'False
            Width           =   3975
            Begin VB.CommandButton cBoton2 
               BackColor       =   &H0000C0C0&
               Caption         =   "&Cancela"
               Height          =   435
               Index           =   1
               Left            =   2100
               Style           =   1  'Graphical
               TabIndex        =   134
               Top             =   1050
               Width           =   1215
            End
            Begin VB.CommandButton cBoton2 
               BackColor       =   &H0000C0C0&
               Caption         =   "&Acepta"
               Height          =   435
               Index           =   0
               Left            =   690
               MaskColor       =   &H0000C0C0&
               Style           =   1  'Graphical
               TabIndex        =   133
               Top             =   1050
               Width           =   1215
            End
            Begin VB.OptionButton cOpc3 
               BackColor       =   &H00800000&
               Caption         =   "FACTURACION"
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
               Height          =   195
               Index           =   1
               Left            =   1950
               TabIndex        =   132
               Top             =   600
               Width           =   1695
            End
            Begin VB.OptionButton cOpc3 
               BackColor       =   &H00800000&
               Caption         =   "MODIFICA"
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
               Height          =   195
               Index           =   0
               Left            =   480
               TabIndex        =   131
               Top             =   600
               Width           =   1275
            End
            Begin VB.Shape Shape1 
               BackColor       =   &H00800000&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H80000006&
               Height          =   1335
               Index           =   4
               Left            =   150
               Shape           =   4  'Rounded Rectangle
               Top             =   270
               Width           =   3735
            End
            Begin VB.Shape Shape1 
               BackColor       =   &H00C0C0C0&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H00FFFFFF&
               BorderStyle     =   6  'Inside Solid
               FillColor       =   &H00C0FFC0&
               FillStyle       =   0  'Solid
               Height          =   1515
               Index           =   3
               Left            =   30
               Shape           =   4  'Rounded Rectangle
               Top             =   180
               Width           =   3945
            End
            Begin VB.Shape Shape1 
               BackColor       =   &H00800000&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H80000006&
               Height          =   1395
               Index           =   2
               Left            =   120
               Shape           =   4  'Rounded Rectangle
               Top             =   240
               Width           =   3765
            End
         End
         Begin VB.Frame Frame5 
            Height          =   585
            Index           =   0
            Left            =   9540
            TabIndex        =   119
            Top             =   6570
            Width           =   2265
            Begin VB.TextBox TReg 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   285
               Left            =   1350
               TabIndex        =   121
               Top             =   210
               Width           =   765
            End
            Begin VB.Label Label3 
               Caption         =   "Total Reg."
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
               Index           =   5
               Left            =   150
               TabIndex        =   120
               Top             =   270
               Width           =   1035
            End
         End
         Begin TrueOleDBGrid70.TDBGrid TDBGrid2 
            Height          =   5355
            Left            =   210
            TabIndex        =   141
            Top             =   240
            Width           =   11205
            _ExtentX        =   19764
            _ExtentY        =   9446
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
      End
      Begin VB.Frame Frame6 
         Height          =   550
         Left            =   -74970
         TabIndex        =   153
         Top             =   3168
         Width           =   11775
         Begin VB.TextBox Text4 
            Height          =   285
            Index           =   2
            Left            =   10440
            MaxLength       =   8
            TabIndex        =   47
            Top             =   180
            Width           =   1215
         End
         Begin VB.TextBox Text4 
            Height          =   285
            Index           =   1
            Left            =   9900
            MaxLength       =   3
            TabIndex        =   46
            Top             =   180
            Width           =   495
         End
         Begin VB.TextBox Text4 
            Height          =   285
            Index           =   0
            Left            =   9480
            MaxLength       =   2
            TabIndex        =   45
            Top             =   180
            Width           =   375
         End
         Begin VB.TextBox Text3 
            Height          =   285
            Left            =   1200
            MaxLength       =   254
            TabIndex        =   44
            Top             =   180
            Visible         =   0   'False
            Width           =   7335
         End
         Begin VB.Label Label8 
            Caption         =   "Referencia"
            Height          =   255
            Left            =   8640
            TabIndex        =   155
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label7 
            Caption         =   "Observacion"
            Height          =   255
            Left            =   120
            TabIndex        =   154
            Top             =   240
            Visible         =   0   'False
            Width           =   975
         End
      End
      Begin VB.Frame Fr2 
         Height          =   885
         Index           =   1
         Left            =   -74970
         TabIndex        =   3
         Top             =   3708
         Width           =   11835
         Begin VB.CommandButton cAyuda 
            Caption         =   "..."
            Height          =   375
            Index           =   3
            Left            =   3660
            TabIndex        =   51
            Top             =   420
            Width           =   285
         End
         Begin MSMask.MaskEdBox MBox2 
            Height          =   375
            Index           =   0
            Left            =   1530
            TabIndex        =   49
            Top             =   420
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   661
            _Version        =   393216
            ClipMode        =   1
            MaxLength       =   6
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MBox2 
            Height          =   375
            Index           =   1
            Left            =   2340
            TabIndex        =   50
            Top             =   420
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   661
            _Version        =   393216
            MaxLength       =   8
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MBox2 
            Height          =   375
            Index           =   2
            Left            =   7710
            TabIndex        =   59
            Top             =   420
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   661
            _Version        =   393216
            BackColor       =   -2147483648
            Enabled         =   0   'False
            MaxLength       =   8
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MBox2 
            Height          =   375
            Index           =   3
            Left            =   8610
            TabIndex        =   52
            Top             =   420
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   661
            _Version        =   393216
            ClipMode        =   1
            MaxLength       =   10
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MBox2 
            Height          =   375
            Index           =   4
            Left            =   9810
            TabIndex        =   53
            Top             =   420
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   661
            _Version        =   393216
            ClipMode        =   1
            MaxLength       =   10
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MBox2 
            Height          =   375
            Index           =   5
            Left            =   10740
            TabIndex        =   54
            Top             =   420
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   661
            _Version        =   393216
            ClipMode        =   1
            MaxLength       =   10
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MBox2 
            Height          =   375
            Index           =   11
            Left            =   90
            TabIndex        =   4
            Top             =   420
            Width           =   585
            _ExtentX        =   1032
            _ExtentY        =   661
            _Version        =   393216
            ClipMode        =   1
            BackColor       =   -2147483644
            Enabled         =   0   'False
            MaxLength       =   6
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MBox2 
            Height          =   375
            Index           =   12
            Left            =   720
            TabIndex        =   48
            Top             =   420
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   661
            _Version        =   393216
            Enabled         =   0   'False
            MaxLength       =   6
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MBox2 
            Height          =   375
            Index           =   13
            Left            =   90
            TabIndex        =   143
            Top             =   420
            Width           =   585
            _ExtentX        =   1032
            _ExtentY        =   661
            _Version        =   393216
            BackColor       =   -2147483648
            Enabled         =   0   'False
            MaxLength       =   6
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MBox2 
            Height          =   375
            Index           =   14
            Left            =   90
            TabIndex        =   144
            Top             =   420
            Width           =   585
            _ExtentX        =   1032
            _ExtentY        =   661
            _Version        =   393216
            ClipMode        =   1
            BackColor       =   -2147483644
            Enabled         =   0   'False
            MaxLength       =   6
            PromptChar      =   "_"
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
            TabIndex        =   142
            Top             =   180
            Width           =   765
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
            TabIndex        =   13
            Top             =   180
            Width           =   1215
         End
         Begin VB.Label deta 
            Alignment       =   2  'Center
            Caption         =   "Descripci�n"
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
            TabIndex        =   12
            Top             =   180
            Width           =   3885
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
            TabIndex        =   11
            Top             =   180
            Width           =   675
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
            TabIndex        =   10
            Top             =   180
            Width           =   1005
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
            TabIndex        =   9
            Top             =   180
            Width           =   735
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
            TabIndex        =   8
            Top             =   180
            Width           =   975
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
            TabIndex        =   7
            Top             =   180
            Width           =   795
         End
         Begin VB.Label Label2 
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   3960
            TabIndex        =   6
            Top             =   420
            Width           =   3675
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
            TabIndex        =   5
            Top             =   180
            Width           =   465
         End
      End
      Begin TrueOleDBGrid70.TDBGrid TDBGrid3 
         Height          =   5220
         Left            =   -74715
         TabIndex        =   166
         Top             =   720
         Width           =   11070
         _ExtentX        =   19526
         _ExtentY        =   9208
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "item"
         Columns(0).DataField=   "item"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Articulo"
         Columns(1).DataField=   "articulo"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "descripcion"
         Columns(2).DataField=   "descripcion"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "unidad"
         Columns(3).DataField=   "unidad"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "saldo"
         Columns(4).DataField=   "saldo"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "cantidad"
         Columns(5).DataField=   "cantidad"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   6
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   13160660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=6"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=900"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=820"
         Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(5)=   "Column(1).Width=2514"
         Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2434"
         Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(9)=   "Column(2).Width=9128"
         Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=9049"
         Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(13)=   "Column(3).Width=1640"
         Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=1561"
         Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(17)=   "Column(4).Width=2064"
         Splits(0)._ColumnProps(18)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(19)=   "Column(4)._WidthInPix=1984"
         Splits(0)._ColumnProps(20)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(21)=   "Column(5).Width=2302"
         Splits(0)._ColumnProps(22)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(5)._WidthInPix=2223"
         Splits(0)._ColumnProps(24)=   "Column(5).Order=6"
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
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=12,.bold=0,.fontsize=780,.italic=0"
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
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=54,.parent=13"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=51,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=52,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=53,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=28,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=58,.parent=13"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=55,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=56,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=57,.parent=17"
         _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=46,.parent=13"
         _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=43,.parent=14"
         _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=44,.parent=15"
         _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=45,.parent=17"
         _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=50,.parent=13"
         _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=47,.parent=14"
         _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=48,.parent=15"
         _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=49,.parent=17"
         _StyleDefs(60)  =   "Named:id=33:Normal"
         _StyleDefs(61)  =   ":id=33,.parent=0"
         _StyleDefs(62)  =   "Named:id=34:Heading"
         _StyleDefs(63)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(64)  =   ":id=34,.wraptext=-1"
         _StyleDefs(65)  =   "Named:id=35:Footing"
         _StyleDefs(66)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(67)  =   "Named:id=36:Selected"
         _StyleDefs(68)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(69)  =   "Named:id=37:Caption"
         _StyleDefs(70)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(71)  =   "Named:id=38:HighlightRow"
         _StyleDefs(72)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(73)  =   "Named:id=39:EvenRow"
         _StyleDefs(74)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(75)  =   "Named:id=40:OddRow"
         _StyleDefs(76)  =   ":id=40,.parent=33"
         _StyleDefs(77)  =   "Named:id=41:RecordSelector"
         _StyleDefs(78)  =   ":id=41,.parent=34"
         _StyleDefs(79)  =   "Named:id=42:FilterBar"
         _StyleDefs(80)  =   ":id=42,.parent=33"
      End
      Begin VB.Label Label4 
         Caption         =   "Label4"
         Height          =   390
         Index           =   2
         Left            =   -69045
         TabIndex        =   160
         Top             =   6165
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Label3"
         Height          =   390
         Index           =   6
         Left            =   -71010
         TabIndex        =   159
         Top             =   6165
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Cantidad"
         Height          =   390
         Index           =   26
         Left            =   -74610
         TabIndex        =   158
         Top             =   6165
         Width           =   975
      End
   End
   Begin Crystal.CrystalReport oCrystalReport 
      Left            =   300
      Top             =   8070
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileUseRptNumberFmt=   -1  'True
      PrintFileUseRptDateFmt=   -1  'True
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
End
Attribute VB_Name = "FrmPedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
                    
Option Explicit

Dim nLongicampo(6) As Integer
Dim rsdeta As New ADODB.Recordset
Dim rsmasivo As New ADODB.Recordset
Dim wCabe(43)

'****** Totales de Pedidos***
Dim Tbruto As Double
Dim Tigv As Double
Dim Tdscto As Double
Dim TSub As Double
Dim TImporte As Double
Dim TNeto As Double
Dim TCant As Double
Dim Flag As Integer
Dim Ctrlgrilla3 As Integer
Dim masivo As Integer
Dim inicio As Integer
Dim Fin As Integer
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
Const W1TXT2 = "El Cliente No Tiene N�mero de R.U.C. en el Maestro"
Const W1TXT3 = "El Cliente Esta Suspendido No Atender"
Const W1TXT4 = "El Cliente Ya No Tiene Credito. No Atender"

Const W1TXT6 = "Codigo del Vendedor No Existe en Tabla de Vendedores"
Const W1TXT7 = "El Codigo del Almacen No Existe en Tabla de Almacenes"

Const W1TXT9 = "El Monto de Otros Gastos debe ser un Valor Positivo"

Const W1TXT12 = "El Descuento General debe ser un Valor Positivo"
Const W1TXT13 = "El Descuento de Promoci�n debe ser un Valor Positivo"
Const W1TXT14 = "El Descuento Pronto Pago debe ser un Valor Positivo"
Const W1TXT17 = "Codigo de la Lista de Precios No Existe"
Const W1TXT18 = "Archivo Maestro de la Lista de Precios No Existe"
Const W1TXT19 = "Codigo del Art�culo No Existe en Maestro de Art�culos "
Const W1TXT20 = "El Codigo del Articulo No Existe en Maestro de Precios"
Const W1TXT21 = "El Codigo del Articulo Ya Existe en el Proceso de Ventas"
Const W1TXT22 = "La Cantidad a Vender debe ser un Valor Mayor que Cero"
Const W1TXT23 = "La Cantidad a Vender es Mayor que el Actual en Almac�n"
Const W1TXT24 = "El Precio de Venta debe de ser un Valor Mayor que Cero"
Const W1TXT25 = "El Descuento por Item debe ser un Valor Positivo"
Const W1TXT28 = "Debe de Ingresar el Nro. de R.U.C. del Cliente"
Const W1TXT30 = "El Importe debe ser mayor a cero"
Const W1TXT31 = ""
Const W1TXT32 = ""
Const W1TXT33 = ""
Private Sub procImprimirguia()
Dim nguia As String
Dim ntabla As String
Dim busca As New dll_apisgen.dll_apis
Dim rb As New ADODB.Recordset
Dim contador As Double
Dim SQL As String

        nguia = g_pedserie & Right("000000000000" & TraeDataSerie("select puntovtadoccorr from vt_puntovtadocumento where documentocodigo='GR' and puntovtadocserie='" & g_pedserie & "' and puntovtacodigo='" & g_ptoventa & "'", Cn), 8)

        Cn.Execute "Update vt_puntovtadocumento " & _
                " set puntovtadoccorr='" & Right("00000000" & Trim(CStr(Val(nguia) + 1)), 8) & "'" & _
                " Where documentocodigo='GR' and puntovtacodigo='" & g_ptoventa & "' and puntovtadocserie='" & g_pedserie & "'"

        Cn.Execute "UPDATE vt_pedido set pedidonrogiarem='" & nguia & "'" & _
                   " Where pedidonumero='" & MBox(1).ClipText & "'"
                   
   '         oCrystalReport.Reset
            
            If cOpc2(0).Value Then
               ntabla = "vt_detallepedido"
             Else
               If cOpc2(1).Value Then
                  ntabla = "vt_detallepedido"
                Else
                  If cOpc2(2).Value Then
                     ntabla = "vt_detallepedido"
                   Else
                     ntabla = g_DetallePuntoVta
                  End If
               End If
            End If
   
        Cn.Execute "delete from gtempfile2filas"
        Set rb = Cn.Execute("select * from gtempfile")
        If rb.RecordCount > 0 Then
           rb.MoveFirst
           contador = 0
           Do Until rb.EOF
              contador = contador + 1
              If contador <= 25 Then
                 SQL = "INSERT INTO gtempfile2filas(item,producto1,descripcion1,cantidad1,importe1)"
                 SQL = SQL & " VALUES ( '" & contador & "','" & RTrim(rb!productocodigo) & "','" & RTrim(rb!productodescripcion) & "','" & rb!detpedcantpedida & "','" & rb!detpedimpbruto & "')"
                 
               Else
                TCant = contador - 25
                 SQL = "UPDATE gtempfile2filas set producto2 ='" & RTrim(rb!productocodigo) & "',"
                 SQL = SQL & " descripcion2='" & RTrim(rb!productodescripcion) & "',"
                 SQL = SQL & "cantidad2='" & rb!detpedcantpedida & "',"
                 SQL = SQL & "importe2= '" & rb!detpedimpbruto & "'"
                 SQL = SQL & " where item = " & TCant & ""
              End If
              Cn.Execute SQL
              rb.MoveNext
           Loop
           rb.Close
           Set rb = Nothing
        End If
   
   
        oCrystalReport.Reset
        oCrystalReport.ReportFileName = RutaRep & "Repguiaimpresa.rpt"
        oCrystalReport.LogOnServer "pdssql.dll", _
              busca.LeerIni(App.Path & "\Camtex.ini", "Bventas", "dserver", ""), _
              busca.LeerIni(App.Path & "\Camtex.ini", "Bventas", "dbase", ""), _
              busca.LeerIni(App.Path & "\Camtex.ini", "Bventas", "duser", ""), _
              busca.LeerIni(App.Path & "\Camtex.ini", "Bventas", "dpass", "")
        oCrystalReport.Connect = _
             "DSN=" & busca.LeerIni(App.Path & "\Camtex.ini", "Bventas", "dserver", "") & ";" & _
             "DSQ=" & busca.LeerIni(App.Path & "\Camtex.ini", "Bventas", "dbase", "") & ";" & _
             "UID=" & busca.LeerIni(App.Path & "\Camtex.ini", "Bventas", "duser", "") & ";" & _
             "PWD=" & busca.LeerIni(App.Path & "\Camtex.ini", "Bventas", "dpass", "")
                
        oCrystalReport.Destination = crptToWindow
        oCrystalReport.WindowState = crptMaximized
        oCrystalReport.DiscardSavedData = True
        With oCrystalReport
               .Formulas(0) = "nro='" & MBox(2) & "'"
               .Formulas(1) = "cliente='" & MBox3(1) & "'"
               .Formulas(2) = "fecha='" & CStr(Day(CDate(MBox(10)))) & "     " & dllgeneral.DESMES(Month(CDate(MBox(10)))) & "                       " & Right(CStr(Year(CDate(MBox(10)))), 1) & "'"
               .Formulas(3) = "direccion='" & MBox3(3) & "'"
               .Formulas(4) = "dni='" & MBox3(2) & "'"
               .Formulas(5) = "opedido='" & MBox(1) & "'"
               .Formulas(6) = "ocompra='" & MBox(17) & "'"
               .Formulas(7) = "guia='" & nguia & "'"
               .Formulas(8) = "distrito='" & MBox3(4).ClipText & "'"
               .Formulas(9) = "destino='" & MBox(19).ClipText & "'"
               Set rb = Cn.Execute("select * from gr_empresa where empresacodigo='" & g_Empresa & "'")
               If rb.RecordCount > 0 Then
                   .Formulas(10) = "partida='" & Escadena(rb!empresadireccion) & "'"
               Else
                   .Formulas(10) = "partida=''"
               End If
               rb.Close
               Set rb = Nothing
               
        End With
        oCrystalReport.Action = 1

End Sub

Private Sub cAyuda_Click(Index As Integer)
  nAyuda = "": nDetalle = ""
  If Index = 0 And Len(Trim(MBox(19))) = 0 Then    'Ayuda de Punto de LLegada
    If dllgeneral.VerificaDatoExistente(Cn, "select * from vt_cliente where clientecodigo='" & Ctr_Ayuda1.xclave & "'") = 1 Then
       Dim gfiltra(1, 2) As String
       gfiltra(1, 1) = "Descripcion": gfiltra(1, 2) = "clientedireccion"
       FrmAyuda.TipoForma = 1
       FrmAyuda.BConexion = Cn
       FrmAyuda.BTabla = "vt_clientedireccion"
       FrmAyuda.Bdata = "0"
       FrmAyuda.BCampos = "Cliedirnumero as Codigo,Cliedirdireccion as Descripcion"
       FrmAyuda.BOrden = "Cliedirnumero"
       FrmAyuda.BCondi = "clientecodigo='" & Ctr_Ayuda1.xclave & "'"
       FrmAyuda.BFiltro = gfiltra
    Else
        nAyuda = "": nDetalle = ""
        MsgBox "No existen Direcciones Anexas...", vbInformation, MsgTitle
        Exit Sub
    End If
  ElseIf Index = 3 Then                             ' Ayuda de Productos
       If Len(Label2) > 0 Then
         SendKeys "{tab}"
         Exit Sub
       End If
       Dim sfiltra(1 To 2, 1 To 2) As String
       sfiltra(1, 1) = "Codigo": sfiltra(1, 2) = "acodigo"
       sfiltra(2, 1) = "Descripcion": sfiltra(2, 2) = "adescri"
       FrmAyuda.TipoForma = 1
       FrmAyuda.BConexion = Cn
       If Combo2.ListCount > 0 Then
             FrmAyuda.BTabla = "[" & cbdatos.DefaultDatabase & "].dbo.maeart inner join [" & _
                            cbdatos.DefaultDatabase & "].dbo.stkart " & _
                            " ON acodigo=stcodigo"
       Else
              FrmAyuda.BTabla = "[" & cbdatos.DefaultDatabase & "].dbo.maeart inner join [" & _
                            cbdatos.DefaultDatabase & "].dbo.stkart " & _
                            " ON acodigo=stcodigo "
       End If
       FrmAyuda.Bdata = "2"
       FrmAyuda.Bdato = Escadena(MBox2(1).Text)
       If modoventa.ctrlinventario = 0 Then
          FrmAyuda.BCampos = "acodigo as Codigo,adescri as Descripcion"
       Else
          FrmAyuda.BCampos = "acodigo as Codigo,adescri as Descripcion,stskdis as Stock"
       End If
       FrmAyuda.BOrden = "adescri"
       If modoventa.ctrlinventario = "1" Then
          FrmAyuda.BCondi = "stalma='" & Ctr_Ayuda3.xclave & "' and stskdis>0"
       Else
          FrmAyuda.BCondi = "stalma='" & Ctr_Ayuda3.xclave & "' "
       End If
       FrmAyuda.BFiltro = sfiltra
   Else
       SendKeys "{tab}"
       Exit Sub
   End If
   FrmAyuda.Show 1
   If Index = 3 Then
       MBox2(1) = Escadena(nAyuda):   Label2 = Escadena(nDetalle)
   ElseIf Index = 0 Then
       MBox(19) = Escadena(nDetalle)
   End If
   nAyuda = "": nDetalle = ""
End Sub

Private Sub cBoton_Click(Index As Integer)
  Dim J As Integer
  If Index = 0 Then
       Fr1(1).Visible = False
       TClie.Value = 0
       Limpiartexto MBox, 2, 9
       MBox(0).Enabled = False:  MBox(1).Enabled = False
       MBox(0).Text = g_ptoventa
       MBox(1) = g_pedserie & Right("000000000000" & TraeDataSerie("select puntovtadoccorr from vt_puntovtadocumento where documentocodigo='" & g_tipoped & "' and puntovtadocserie='" & g_pedserie & "' and puntovtacodigo='" & g_ptoventa & "'", Cn), 8)
       MBox(2) = g_facserie & Right("000000000000" & TraeDataSerie("select puntovtadoccorr from vt_puntovtadocumento where documentocodigo='" & g_tipofac & "' and puntovtadocserie='" & g_facserie & "' and puntovtacodigo='" & g_ptoventa & "'", Cn), 8)
       MBox(3) = g_bolserie & Right("000000000000" & TraeDataSerie("select puntovtadoccorr from vt_puntovtadocumento where documentocodigo='" & g_tipobol & "' and puntovtadocserie='" & g_bolserie & "' and puntovtacodigo='" & g_ptoventa & "'", Cn), 8)
       MBox(4) = g_guiaserie & Right("000000000000" & TraeDataSerie("select puntovtadoccorr from vt_puntovtadocumento where documentocodigo='" & g_tipoguia & "' and puntovtadocserie='" & g_guiaserie & "' and puntovtacodigo='" & g_ptoventa & "'", Cn), 8)
       MBox(5) = Numero(0): MBox(6) = Numero(0): MBox(7) = Numero(0): MBox(8) = Numero(TraeTipoCambio(Date, Cn))
       MBox(9) = Escadena(parametro.mensaje)
       MBox(19) = ""
       MBox(10) = Format(Date, "dd/mm/yyyy")
       MBox(13) = Numero(0)
       MBox(15) = Numero(0)
       MBox(16) = 0: MBox(17) = 0: MBox(18) = 0
       For J = 0 To 5
          MBox2(J) = ""
       Next J
       Set rsdeta = Nothing
       
       CargaGrilla

     'Se activa los parametros deventa
       Combo1.ListIndex = VerificaCombo(Combo1, parametro.moneda)     'moneda
       Combo2.ListIndex = VerificaCombo(Combo2, parametro.listapre)   'listaprecios
       'Combo2.Enabled = False
       MBox(8) = Numero(parametro.tipocambio)                         'tipo de cambio
       Ctr_Ayuda3.xclave = Escadena(parametro.almacen)                'almacen
       Call Ctr_Ayuda3.Ejecutar
'       If Len(Trim(modoventa.almacenes)) > 0 Then
'          Ctr_Ayuda3.Filtro = "almacencodigo in (" & modoventa.almacenes & ")"
'       End If

       
       MBox(13).Enabled = IIf(parametro.comivende = "F", False, True)                     'comision de vendedor
       
      'Se activa los parametros de punto de venta
       MBox(2).Enabled = IIf(punto.nrofactura = "1" And punto.ventaauto = "0", True, False)
       MBox(3).Enabled = IIf(punto.nroboleta = "1" And punto.ventaauto = "0", True, False)
       MBox(4).Enabled = IIf(punto.nroguia = "1" And punto.ventaauto = "0", True, False)
       
     'Activamos el Tab
       Activa 1
       SSTab2.TabEnabled(2) = False
       SSTab2.Tab = 0
       MBox(5).SetFocus

  ElseIf Index = 1 Then
      Fr1(1).Visible = False
  End If
End Sub

Private Sub cBoton2_Click(Index As Integer)
    If Index = 0 Then
        cOpc(0).Value = False
        cOpc(1).Value = False
    Else
        Fr5.Visible = False
        Exit Sub
    End If
    Fr5.Visible = False
    Carga_Pedido
    Activa 1
    MBox(0).Enabled = False
    MBox(1).Enabled = False
    MBox(5).SetFocus
    g_TipoMovi = 2

End Sub



Private Sub cmdBotones_Click(Index As Integer)
   Dim asql As String
   Dim acmd As New ADODB.Command
   Dim J, nl As Integer
   Dim nflag As Integer
   
   On Error GoTo vererror
   'On Error Resume Next
   
   nflag = 0
   Select Case Index
    Case 0
        Fr1(1).Visible = True
        Limpiartexto MBox2, 6, 10
        Fr1(0).Enabled = True
        Fr2(0).Enabled = True
        Fr3(0).Enabled = True
        TClie.Enabled = True
        Text3 = "": Text4(0) = "": Text4(1) = "": Text4(2) = ""
        Call CargarModo
        cOpc(0).Value = False: cOpc(1).Value = False: cOpc2(0).Value = False
        cOpc2(1).Value = False: cOpc2(2).Value = False
        cOpc3(0).Value = False: cOpc3(1).Value = False
        cOpc(0).SetFocus
        g_TipoMovi = 1
        masivo = 0
'        rsmasivo.Reset
        
    Case 1
       If TDBGrid2.Row >= 0 Then
          Fr1(0).Enabled = True
          Fr2(0).Enabled = True
          Fr3(0).Enabled = True
          TClie.Enabled = True
          Limpiartexto MBox2, 6, 10
          cOpc(0).Value = False: cOpc(1).Value = False: cOpc2(0).Value = False
          cOpc2(1).Value = False: cOpc2(2).Value = False
          cOpc3(0).Value = False: cOpc3(1).Value = False
          Text3 = "": Text4(0) = "": Text4(1) = "": Text4(2) = ""
          Fr5.Visible = True
          cOpc3(0).SetFocus
       End If
    Case 2
       If TDBGrid2.Row >= 0 Then
        asql = "pedidonumero='" & TDBGrid2.Columns(0).Text & "'"
        If dllgeneral.EliminaReg(Cn, g_DetallePuntoVta, asql) = 1 Then
            Cn.Execute "Delete From " & g_PedidoPuntoVta & " where " & asql
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
        If dllgeneral.VerificaDatoExistente(Cn, "select * from vt_cliente where clientecodigo='" & MBox3(0) & "' and clientesuspendido='1'") = 1 And MBox3(0) <> g_Eventual Then
           MsgBox W1TXT3, vbInformation, MsgTitle
           Exit Sub
        End If
        If dllgeneral.VerificaDatoExistente(Cn, "select * from vt_cliente where clientecodigo='" & MBox3(0) & "' and ((clientelimitecreddolar-clientesaldodolares)*" & MBox(8) & "+ (clientelimitecredsoles-clientesaldosoles))-" & TNeto & " <=0") = 1 And MBox3(0) <> g_Eventual Then
           MsgBox W1TXT4, vbInformation, MsgTitle
           Exit Sub
        End If
        If Len(Trim(Text4(0))) > 0 Then
            If dllgeneral.VerificaDatoExistente(Cn, "select * from vt_pedido where pedidotipofac='" & Text4(0) & "' and pedidonrofact='" & Trim(Text4(1)) & Trim(Text4(2)) & "'") = 0 Then
               MsgBox "No existe documento...Verifique!!!", vbInformation, "AVISO"
               Exit Sub
            End If
        End If
        
       If cOpc(0).Value Or cOpc3(0).Value Then
          nflag = 1
          Cn.BeginTrans
          If GrabarData() = 1 Then
            Cn.CommitTrans
            nflag = 0
            g_TipoMovi = 0
            If modoventa.emitehoja = "1" Then
               nl = IIf(modoventa.copiashoja > 0, modoventa.copiashoja, 0)
               If nl > 0 Then
                   For J = 1 To nl
                      Call DocImprimir
                   Next J
               End If
            End If
            Activa 2
            Listado
            Exit Sub
          Else
             Cn.RollbackTrans
             nflag = 0
             g_TipoMovi = 0
             Activa 2
             Exit Sub
          End If
       Else
          If cOpc(1).Value Or cOpc3(1).Value Then
                Fr4.Visible = True
                cOpc2(0).Value = Escadena(IIf(modoventa.documento <> g_tipofac, False, True))
                cOpc2(1).Value = Escadena(IIf(modoventa.documento <> g_tipobol, False, True))
                cOpc2(2).Value = Escadena(IIf(modoventa.documento <> g_tipoguia, False, True))
                Exit Sub
           Else
                g_TipoMovi = 0
                Activa 2
                Exit Sub
           End If
       End If
       g_TipoMovi = 0
    Case 12
       Activa 2
       g_TipoMovi = 0
   End Select
   
vererror:
    If Err Then
       If nflag = 1 Then
         Cn.RollbackTrans
       End If
       MsgBox "Comunicarse con Sistemas ...!!" & Chr(13) & Chr(10) & Err.Number & "-" & Err.Description
       Err = 0
       Exit Sub
       Resume
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

Private Sub Cmdgrabamasivo_Click()
    Call grabamasivo
    Call dllgeneral.ActivaTab(1, 2, SSTab1)
End Sub

Private Sub Cmdsalirmasivo_Click()
   Call dllgeneral.ActivaTab(1, 2, SSTab1)
End Sub

Private Sub Combo1_Click()
   MBox(8) = TraeDataSerie("select * from ct_tipocambio where tipocambiofecha=GETDATE()", Cn)
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
  Seguir Combo1, KeyAscii
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
  Seguir Combo2, KeyAscii
End Sub

Private Sub Combo3_Click()
  If Combo3.ListCount >= 0 Then
     Call CargarModo
'     If Len(Trim(modoventa.almacenes)) > 0 Then
'          Ctr_Ayuda3.Filtro = "almacencodigo in (" & modoventa.almacenes & ")"
'          'Ctr_Ayuda3.Ejecutar
'      End If
  End If
End Sub

Private Sub Combo3_GotFocus()
   If Combo3.ListCount - 1 <= 0 Then
      Call dllgeneral.llenacombo(Combo3, "select modovtacodigo,modovtadescripcion from vt_modoventa", Cn)
      Exit Sub
   End If

End Sub


Private Sub Combo3_KeyPress(KeyAscii As Integer)
  Call Combo3_Click
  Seguir Combo3, KeyAscii
End Sub

Private Sub Combo4_GotFocus()
   If Combo4.ListCount - 1 <= 0 Then
       Call dllgeneral.llenacombo(Combo4, "select formapagocodigo,formapagodescripcion from vt_formapago", Cn)
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

Private Sub Command1_Click()
 Call grabamasivo
 Call dllgeneral.ActivaTab(1, 1, SSTab2)
End Sub

Private Sub Command2_Click()
 Call dllgeneral.ActivaTab(1, 1, SSTab2)
   
End Sub

Private Sub cOpc_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = 13 Then
    cBoton(0).SetFocus
  End If
End Sub

Private Sub cSeleccion_Click(Index As Integer)
  Dim nArchi As String
  Dim rsel As New ADODB.Recordset
  Dim nl As Integer
  Dim numpedidos As Integer
  Dim contador As Integer
  Dim residuo As Integer
  Dim J As Integer
  Dim nflag As Integer
  
  'On Error GoTo nerror
  
  Cn.Execute "delete from gtempfile"
  Cn.Execute "delete from tempfile"
  
  nflag = 0
  If Index = 0 Then
    If cOpc2(0).Value And dllgeneral.VerificaDatoExistente(Cn, "select * from vt_cliente where clientecodigo='" & Ctr_Ayuda1.xclave & "' and len(ltrim(clienteruc))=11") = 0 Then
        MsgBox "El cliente no tiene ruc valido....Verifique!!!", vbInformation, MsgTitle
        Exit Sub
    Else
        Set rsel = Cn.Execute("select * from vt_cliente where clientecodigo='" & Ctr_Ayuda1.xclave & "' and clientecodigo<>'" & Left(g_Eventual, Len(Trim(Ctr_Ayuda1.xclave))) & "'")
        If rsel.RecordCount > 0 Then
           MBox3(0) = Escadena(rsel!clientecodigo)
           MBox3(1) = Escadena(rsel!clienterazonsocial)
           MBox3(2) = Escadena(rsel!clienteruc)
           MBox3(3) = Escadena(rsel!clientedireccion)
           MBox3(4) = Escadena(rsel!clientedistrito)
        End If
        rsel.Close
        Set rsel = Nothing
    End If
    cSeleccion(0).Enabled = False
    nflag = 1
    numpedidos = rsdeta.RecordCount
    contador = Int(numpedidos / 100)
    residuo = numpedidos Mod 100
    If residuo > 0 Then contador = contador + 1
    numpedidos = 0
    rsdeta.MoveFirst
    Do While numpedidos < contador
        numpedidos = numpedidos + 1
        inicio = (numpedidos - 1) * 100 + 1
        If numpedidos * 100 > rsdeta.RecordCount Then
           Fin = rsdeta.RecordCount
         Else
           Fin = numpedidos * 100
        End If
        Cn.BeginTrans
        If (cOpc2(0).Value Or cOpc2(1).Value Or cOpc2(2).Value) And (cOpc2(0).Enabled Or cOpc2(1).Enabled Or cOpc2(2).Enabled) Then
          If GrabarData() = 1 Then
             Cn.CommitTrans
             nflag = 0
             If modoventa.emitefact = "1" Or modoventa.emiteguia = "1" Then
                nl = IIf(modoventa.copiasbol > 0, modoventa.copiasbol, 0)
                If nl <= 0 Then
                   nl = IIf(modoventa.copiasfac > 0, modoventa.copiasfac, 0)
                End If
                If nl > 0 Then
                    For J = 1 To nl
                       Call DocImprimir
                    Next J
                End If
             End If
             Listado
          Else
             Cn.RollbackTrans
             nflag = 0
          End If
        Else
          Cn.RollbackTrans
          nflag = 0
          cSeleccion(0).Enabled = True
          MsgBox "Seleccione un tipo de Documento para Grabar...!!!", vbInformation, MsgTitle
          Exit Sub
        End If
    Loop
  End If
  cSeleccion(0).Enabled = True
  Fr4.Visible = False
  Activa 2
  
nerror:
   If Err Then
      If nflag = 1 Then
         Cn.RollbackTrans
      End If
      MsgBox "Comunicarse con Sistemas ...!!" & Chr(13) & Chr(10) & Err.Number & "-" & Err.Description
      Exit Sub
   End If
End Sub

Private Sub Ctr_Ayuda1_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
    Dim acliente As New ADODB.Recordset
    
    MBox3(0) = Trim(ColecCampos.Item(0))
    MBox3(1) = Trim(ColecCampos.Item(1))
    MBox3(2) = Trim(ColecCampos.Item(2))
    MBox(19) = Trim(ColecCampos.Item(3))
    MBox3(3) = Trim(ColecCampos.Item(3))
    MBox3(4) = Trim(ColecCampos.Item(4))
    Text5 = Trim(ColecCampos.Item(2))
    
    If IsNull(ColecCampos.Item(10)) Or Len(Trim(ColecCampos.Item(10))) = 0 Then
       text1 = Numero(0)
       Text2 = Numero(0)
    Else
       text1 = Numero(CDbl(Trim(ColecCampos.Item(10))))
       Text2 = Numero(CDbl(Trim(ColecCampos.Item(10))) * 100)
    End If
    
    lcred(0) = Numero(0)
    lcred(1) = Numero(0)

    Set acliente = Cn.Execute("select * from vt_cliente where clientecodigo='" & Ctr_Ayuda1.xclave & "'")
    If acliente.RecordCount > 0 Then
       Combo6.ListIndex = VerificaCombo(Combo6, acliente!clientetipopersona)
       Combo7.ListIndex = VerificaCombo(Combo7, acliente!clientetipopais)
       Combo8.ListIndex = VerificaCombo(Combo8, IIf(acliente!clientemultidireccion = 1, "S", "N"))
       lcred(0) = Numero(acliente!clientesaldodolares)
       lcred(1) = Numero(acliente!clientelimitecreddolar)
    End If
    acliente.Close
    Set acliente = Nothing

End Sub

Private Sub Ctr_Ayuda3_GotFocus()
    If Len(Trim(modoventa.almacenes)) > 0 Then
       If Len(Trim(modoventa.almacenes)) <= 2 Then
           Ctr_Ayuda3.Filtro = " almacencodigo like (" & Trim(modoventa.almacenes) & ")"
        ElseIf Len(Trim(modoventa.almacenes)) > 2 Then
           
           Ctr_Ayuda3.Filtro = " almacencodigo in (" & Trim(modoventa.almacenes) & ")"
        Else
           Ctr_Ayuda3.Filtro = " almacencodigo like '%'"
        End If
       Ctr_Ayuda3.Ejecutar
    Else
       Ctr_Ayuda3.Filtro = " almacencodigo like '%'"
       Ctr_Ayuda3.Ejecutar
    End If
End Sub

Private Sub Form_Activate()
  Listado

End Sub

Private Sub Form_Load()
   Call configuramasivo
   MostrarForm Me, "C"
   
   Flag = 0
   'Call dllgeneral.ActivaTab(0, 1, SSTab1)
   Call dllgeneral.ActivaTab(0, 1, SSTab1)
   
   nLongicampo(1) = 1000:  nLongicampo(2) = 1200:   nLongicampo(3) = 6300:   nLongicampo(4) = 600:  nLongicampo(5) = 1200
   
   MBox(1).Enabled = False: Label2 = ""
   Call Cargacombo
   Listado
   Call dllgeneral.ActivaTab(0, 2, SSTab1)
  
   
End Sub
Private Function configuramasivo()
   Set rsmasivo = Nothing
   Call rsmasivo.Fields.Append("item", adVarChar, 5)
   Call rsmasivo.Fields.Append("Articulo", adVarChar, 20)
   Call rsmasivo.Fields.Append("descripcion", adVarChar, 30)
   Call rsmasivo.Fields.Append("unidad", adVarChar, 10)
   Call rsmasivo.Fields.Append("saldo", adVarChar, 30)
   Call rsmasivo.Fields.Append("cantidad", adVarChar, 20)
   rsmasivo.Open
   masivo = 0
   TDBGrid3.Columns(0).AllowFocus = False
   TDBGrid3.Columns(1).AllowFocus = False
   TDBGrid3.Columns(2).AllowFocus = False
   TDBGrid3.Columns(3).AllowFocus = False
   TDBGrid3.Columns(4).AllowFocus = True
   'TDBGrid3.AllowUpdate = True
   TDBGrid3.Columns(4).NumberFormat = "###,##0.0000"
   Set TDBGrid3.DataSource = rsmasivo
End Function
Private Function grabamasivo()
Dim rsk As New ADODB.Recordset
   If rsmasivo.RecordCount > 0 And masivo = 1 Then
      rsmasivo.MoveFirst
      Do Until rsmasivo.EOF
         If rsmasivo!cantidad > 0 Then
            rsdeta.AddNew
            rsdeta.Fields(0) = rsmasivo!Item
            rsdeta.Fields(1) = rsmasivo!articulo
            rsdeta.Fields(2) = rsmasivo!descripcion
            rsdeta.Fields(3) = rsmasivo!unidad
            rsdeta.Fields(4) = rsmasivo!cantidad
            MBox2(0) = rsmasivo!cantidad
            MBox2(2) = rsmasivo!unidad
            MBox2(1) = rsmasivo!articulo
            MBox2(4) = 0
            MBox2(5) = 0
            Set rsk = Cn.Execute("select * from listapre" & Trim(Combo2.Text) & " where productocodigo='" & rsmasivo!articulo & "' and almacencodigo='" & Trim(Ctr_Ayuda3.xclave) & "'")
           If rsk.RecordCount > 0 Then
              rsdeta.Fields(12) = rsk.Fields("productoprecvta")
              MBox2(3) = rsdeta.Fields(12)
            Else
              rsdeta.Fields(12) = 0
              MBox2(3) = 0
           End If
           If parametro.tieneigv = "1" Then
              rsdeta.Fields(5) = (MBox2(3) / (1 + parametro.igv))
'             rsdeta.Fields(12) = MBox2(3).Tag
            Else
              If modoventa.impuestos = "1" Then
                 rsdeta.Fields(5) = (MBox2(3) / (1 + parametro.igv))
                 rsdeta.Fields(12) = MBox2(3).Tag
               Else
                 rsdeta.Fields(5) = MBox2(3).Text
                 rsdeta.Fields(12) = MBox2(3).Tag
              End If
           End If
           rsdeta.Fields(6) = Numero(MBox2(4))
           rsdeta.Fields(7) = Numero(MBox2(0) * MBox2(3))   ' IIf(parametro.tieneigv = "1", (MBox2(3) / (1 + (parametro.igv / 100))), MBox2(3)))
           rsdeta.Fields(8) = rsmasivo!cantidad
           rsdeta.Fields(9) = IIf(Len(Trim(MBox2(12))) = 0, 0, Format(MBox2(12), "##,###,##0"))
           rsdeta.Fields(10) = Numero(MBox2(13))
           rsdeta.Fields(11) = IIf(IsNull(MBox2(14)) Or Len(Trim(MBox2(14))) = 0, 0, MBox2(14))
           rsk.Close
           Set rsk = Nothing
           rsdeta.Update
         End If
         rsmasivo.MoveNext
      Loop
      If rsmasivo.RecordCount > 0 Then
         rsmasivo.MoveFirst
         Do Until rsmasivo.EOF
            rsmasivo.Delete
            rsmasivo.MoveNext
         Loop
      End If
      masivo = 2
   End If
   inicio = 1
   Fin = rsdeta.RecordCount
   rsdeta.MoveFirst
   Totales
End Function
Private Function loadmasivo()
   Dim rsgrid3 As New ADODB.Recordset
   Dim SQL As String
   Dim wposi As Integer
   SQL = " select articulo=stcodigo,descripcion=adescri,unidad=aunidad,saldo=stskdis,cantidad=0 "
   SQL = SQL & " from stkart,maeart where acodigo=stcodigo and stalma='" & Trim(Ctr_Ayuda3.xclave) & "' and alinea = '09' order by stcodigo"
   Set rsgrid3 = Cn.Execute(SQL)
   wposi = 0
   Text10 = 0
   If rsgrid3.RecordCount > 0 And masivo = 0 Then
      rsgrid3.MoveFirst
      Do Until rsgrid3.EOF
         rsmasivo.AddNew
         wposi = wposi + 1
         rsmasivo.Fields(0) = wposi
         rsmasivo.Fields(1) = rsgrid3!articulo
         rsmasivo.Fields(2) = Left(rsgrid3!descripcion, 30)
         rsmasivo.Fields(3) = rsgrid3!unidad
         rsmasivo.Fields(4) = rsgrid3!saldo
         rsmasivo.Fields(5) = rsgrid3!cantidad
         rsgrid3.MoveNext
     Loop
  End If
  rsmasivo.MoveFirst
  TDBGrid3.SetFocus
  TDBGrid3.Refresh
  rsgrid3.Close
  Set rsgrid3 = Nothing
  masivo = 1

End Function

Public Function Cargacombo()
   Dim J As Integer
   Dim nsql As String
   Dim rsk As New ADODB.Recordset
   
   CargaGrilla
   MBox2(11) = rsdeta.RecordCount
   If MBox2(11) > modoventa.nroitem Then
      MsgBox "No se puede Ingresar mas Items...Verifique!!!", vbInformation, MsgTitle
      Exit Function
   End If
  
  
   Call dllgeneral.llenacombo(Combo1, "select monedacodigo,monedadescripcion from gr_moneda order by monedacodigo", Cn)
   'If Combo1.ListCount - 1 >= 0 Then
       Combo1.ListIndex = 0
   'End If
   
   Combo2.Clear
   Set rsk = Cn.Execute("select right(name,1) from sysobjects where name like 'listapre%'")
   If rsk.RecordCount > 0 Then
      rsk.MoveFirst
      Do Until rsk.EOF
        Combo2.AddItem rsk.Fields(0)
        rsk.MoveNext
      Loop
   Else
     Combo2.AddItem "*ninguno"
   End If
   rsk.Close
   Set rsk = Nothing
'     Combo2.AddItem Trim(Str(J))
'   Next J
   Combo2.ListIndex = 0
   
   Call dllgeneral.llenacombo(Combo3, "select modovtacodigo,modovtadescripcion from vt_modoventa", Cn)
   If Combo3.ListCount - 1 >= 0 Then
     Combo3.ListIndex = 0
   End If
   
   Call dllgeneral.llenacombo(Combo4, "select formapagocodigo,formapagodescripcion from vt_formapago", Cn)
   If Combo4.ListCount - 1 >= 0 Then
       Combo4.ListIndex = 0
   End If
   
   
   Call CargarTipo(Combo5, 3)
   
   Call CargarTipo(Combo6, 4)
   
   Call CargarTipo(Combo7, 5)
   
   Call CargarTipo(Combo8, 3)
   
   
   Call Ctr_Ayuda1.conexion(Cn)
   Call Ctr_Ayuda2.conexion(Cn)
   Call Ctr_Ayuda3.conexion(Cn)
   
End Function

Public Function CargaGrilla()

   Call rsdeta.Fields.Append("Item", adInteger)
   Call rsdeta.Fields.Append("Codigo", adVarChar, 20)
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
   Call rsdeta.Fields.Append("PrecioLista", adDouble)
   
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
      .Columns(2).Width = 4000
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
      .Columns(6).NumberFormat = "###,##0.0000"
      .Columns(7).NumberFormat = "###,##0.0000"
      .Columns(8).NumberFormat = "###,##0.0000"
      .Columns(9).Width = 800
      .Columns(9).Caption = "Cant.Ref"
      .Columns(9).NumberFormat = "###,##0"
      .Columns(10).Width = 600
      .Columns(10).Caption = "Factor"
      .Columns(10).NumberFormat = "###,##0.0000"
      .Columns(11).Width = 0
      .Columns(11).NumberFormat = "###,##0.0000"
      .Columns(12).Visible = True
      .Columns(11).Width = 100
   End With
   TDBGrid1.Refresh
End Function
Public Function Listado()
  Call dllgeneral.ListarEnTDBGRID(Cn, g_PedidoPuntoVta, TDBGrid2, "pedidonumero as Pedido,pedidofecha as Fecha,pedidonotaped as Cotizacion,pedidomensaje as Descripcion", "pedidofecha,pedidonumero", nLongicampo)
  TReg.Text = Format(TDBGrid2.ApproxCount, "#########0")
  With TDBGrid2
    .Columns(0).Width = 1200
    .Columns(1).Width = 1200
    .Columns(2).Width = 1200
    .Columns(3).Width = 6500
    .AllowUpdate = False
    .Refresh
  End With
  
  
End Function

Private Sub Form_Unload(Cancel As Integer)
  Set rsdeta = Nothing
End Sub

Private Sub MBox_GotFocus(Index As Integer)
  Call dllgeneral.Enfoquetexto(MBox(Index))
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
  ElseIf KeyCode = 13 And (Index = 19 And Len(Trim(MBox(19))) > 0) Then
        MBox(19) = Escadena(UCase(Trim(MBox(19).ClipText)))
        If IsNull(MBox(19)) Or Len(Trim(MBox(19))) = 0 Then
           MsgBox "Falta Punto de LLegada", vbInformation, MsgTitle
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
        If dllgeneral.VerificaDatoExistente(Cn, "select * from vt_cliente where clientecodigo='" & MBox3(0) & "' and clientesuspendido='1'") = 1 And MBox3(0) <> g_Eventual Then
           MsgBox W1TXT3, vbInformation, MsgTitle
           Exit Sub
        End If
        If dllgeneral.VerificaDatoExistente(Cn, "select * from vt_cliente where clientecodigo='" & MBox3(0) & "' and ((clientelimitecreddolar-clientesaldodolares)*" & MBox(8) & "+ (clientelimitecredsoles-clientesaldosoles)) <=0") = 1 And MBox3(0) <> g_Eventual Then
           MsgBox W1TXT4, vbInformation, MsgTitle
           Exit Sub
        End If
        Fr1(0).Enabled = False
        Fr2(0).Enabled = False
        Fr3(0).Enabled = False
        TClie.Enabled = False
        Call CargarModo
        If Text3.Visible = True Then
           Text3.SetFocus ' "{tab}"
         Else
           Text4(0).SetFocus
        End If
  End If
End Sub


Private Sub MBox_LostFocus(Index As Integer)
  On Error Resume Next
  Select Case Index
   Case 5, 6, 7, 8, 13, 15
      If Not dllgeneral.ValidaCadena(MBox(Index), "N") Then
         MsgBox Msg29, vbInformation, "AVISO"
         Call dllgeneral.Enfoquetexto(MBox(Index))
         Exit Sub
      End If
      MBox(Index) = Format(MBox(Index), "##,##0.0000")
   Case 10
      If Not dllgeneral.ValidaCadena(MBox(Index), "F") Then
         MsgBox "Fecha No Valida", vbInformation, "AVISO"
         Call dllgeneral.Enfoquetexto(MBox(Index))
         Exit Sub
      End If
   Case 16, 17
      If Not dllgeneral.ValidaCadena(MBox(Index), "D") Then
         MsgBox Msg29, vbInformation, "AVISO"
         Call dllgeneral.Enfoquetexto(MBox(Index))
         Exit Sub
      End If
      MBox(Index) = Right("000000000000" & MBox(Index), MBox(Index).MaxLength)
   Case 19
      MBox(19) = Escadena(UCase(Trim(MBox(19).ClipText)))
   Case 18
      If Not dllgeneral.ValidaCadena(MBox(Index), "D") Then
         MsgBox Msg29, vbInformation, "AVISO"
         Call dllgeneral.Enfoquetexto(MBox(Index))
         Exit Sub
      End If
      MBox(Index) = Format(MBox(Index), "####0")
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
  If Index = 3 Then
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
    ElseIf Index = 1 Then
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


Private Sub MBox2_LostFocus(Index As Integer)
  Dim nregi As Long
  Dim wposi, posi As Integer
  Dim ntabla As String
  Dim wflag As Integer
  Dim rsk As New ADODB.Recordset
  
  On Error Resume Next
  
  Select Case Index
   Case 0
      If Not (dllgeneral.ValidaCadena(MBox2(Index), "N") Or IsNumeric(MBox2(Index))) Then
         MsgBox Msg29, vbInformation, "AVISO"
         Call dllgeneral.Enfoquetexto(MBox2(Index))
         Exit Sub
      End If
   Case 1
      'ntabla = IIf(Combo2.ListCount > 0, "listapre" & dllgeneral.ComboDato(Combo2.Text), "vt_producto")
      'If dllgeneral.VerificaDatoExistente(cn, "select * from " & ntabla & " where productocodigo='" & MBox2(Index).Text & "' and almacencodigo='" & Ctr_Ayuda3.xclave & "'") = 0 And Len(Trim(MBox2(Index))) > 0 Then
      If dllgeneral.VerificaDatoExistente(cbdatos, "select * from stkart where stcodigo='" & MBox2(Index).Text & "' and stalma='" & Ctr_Ayuda3.xclave & "'") = 0 And Len(Trim(MBox2(Index))) > 0 Then
          Call cAyuda_Click(3)
          MBox2(1).MaxLength = 20
         Exit Sub
      Else
        wflag = verificaproducto()
        If wflag = 1 Then
            Label2 = ""
            MsgBox "Ya ingreso el producto...Verifique!!!", vbInformation, MsgTitle
            MBox2(1).SetFocus
            Exit Sub
         End If
            
      End If
   Case 3, 4, 5
      If Index = 3 And dllgeneral.ComboDato(Combo5.Text) = "N" Then
          Call TraerProducto
      End If
      If Not dllgeneral.ValidaCadena(MBox2(Index), "N") And Len(Trim(MBox2(Index))) <> 0 Then
         MsgBox Msg29, vbInformation, "AVISO"
         Call dllgeneral.Enfoquetexto(MBox2(Index))
         Exit Sub
      End If
      If Not (dllgeneral.ValidaCadena(MBox2(0), "N") Or IsNumeric(MBox2(0))) Then
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
         MBox2(Index) = Format(MBox2(Index), "######0.0000")  ' Numero(MBox2(Index))
       Else
         MBox2(Index) = Format(MBox2(Index), "######0.0000")
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
        rsdeta.Fields(1) = Trim(Escadena(MBox2(1)))
        rsdeta.Fields(2) = Left(Escadena(Label2) & Space(40), 40)
        rsdeta.Fields(3) = Trim(MBox2(2))
        rsdeta.Fields(4) = Escadena(MBox2(0))
        If parametro.tieneigv = "1" Then
           rsdeta.Fields(5) = (MBox2(3) / (1 + parametro.igv))
           rsdeta.Fields(12) = MBox2(3).Tag
        Else
           If modoventa.impuestos = "1" Then
              rsdeta.Fields(5) = (MBox2(3) / (1 + parametro.igv))
              rsdeta.Fields(12) = MBox2(3).Tag
           Else
              rsdeta.Fields(5) = MBox2(3).Text
              rsdeta.Fields(12) = MBox2(3).Tag
           End If
        End If
        rsdeta.Fields(6) = Numero(MBox2(4))
        rsdeta.Fields(7) = Numero(MBox2(0) * MBox2(3))   ' IIf(parametro.tieneigv = "1", (MBox2(3) / (1 + (parametro.igv / 100))), MBox2(3)))
        rsdeta.Fields(8) = Numero(MBox2(5))
        rsdeta.Fields(9) = IIf(Len(Trim(MBox2(12))) = 0, 0, Format(MBox2(12), "##,###,##0"))
        rsdeta.Fields(10) = Numero(MBox2(13))
        rsdeta.Fields(11) = IIf(IsNull(MBox2(14)) Or Len(Trim(MBox2(14))) = 0, 0, MBox2(14))
        Set rsk = Cn.Execute("select * from listapre" & Trim(Combo2.Text) & " where productocodigo='" & Trim(Escadena(MBox2(1))) & "' and almacencodigo='" & Trim(Ctr_Ayuda2.xclave) & "'")
        If rsk.RecordCount > 0 Then
           rsdeta.Fields(12) = rsk.Fields("productoprecvta")
        Else
           rsdeta.Fields(12) = 0
        End If
        rsk.Close
        Set rsk = Nothing
        
        rsdeta.Update
        Label2 = ""
'        TDBGrid1.Row = rsdeta.RecordCount - 1
        
        ConfigGrid
        Totales
        MBox2(11) = wposi + 1
        If MBox2(12).Enabled = True Then
          MBox2(12).SetFocus
        Else
          MBox2(0).SetFocus
        End If
        Flag = 0
        Exit Sub
    End If
  End Select

End Sub

Private Sub MBox3_KeyPress(Index As Integer, KeyAscii As Integer)
   Seguir MBox3(Index), KeyAscii
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
  If SSTab1.Tab = 2 And Chkmasivo = 0 Then
     MBox2(0).SetFocus
  ElseIf SSTab1.Tab = 1 And Chkmasivo = 0 Then
     If MBox(0).Enabled = True Then
        MBox(5).SetFocus
     Else
        MBox(5).SetFocus
     End If
  End If
End Sub

Public Function Totales()
  Dim J As Double
  Dim Previo As Double
  Dim dct02, dct03, dct04, dct05, dct06 As Double
  
  Tbruto = 0: Tigv = 0: Tdscto = 0: TNeto = 0: TCant = 0
  TImporte = 0: TSub = 0
  '--Totales de Descuentos
  DTGlobal = 0: DTCliente = 0: DTPPago = 0: DTOficina = 0: DTItem = 0
  DTLinea = 0: DTPromo = 0
 
  
  If rsdeta.RecordCount > 0 Then
     For J = inicio To Fin
       'IMPORTE DE MONTO BRUTO SIN IGV, ES DECIR PRECIO X CANTIDAD
       
       Tbruto = Tbruto + (rsdeta.Fields(4) * rsdeta.Fields(5))
       TCant = TCant + rsdeta.Fields(4)
       TImporte = (rsdeta.Fields(4) * rsdeta.Fields(5))
       
       'DESCUENTO DE CIA O EMPRESA
       'If parametro.tienedscto = "1" Then
       '     dct06 = TImporte * (1 + parametro.descuento)
       'Else
       '    dct06 = 0
       'End If
       If IsNull(text1) Or Len(Trim(text1)) = 0 Then
           dct06 = 0
       Else
          'dct06 = TImporte * (1 + parametro.descuento)
          'dct06 = TImporte * (CDbl(Text1))
          dct06 = 0
       End If
       
       DTCliente = DTCliente + dct06
       
       'DESCUENTO POR ITEM
       dct02 = 0
       dct02 = (TImporte * (rsdeta.Fields(6) / 100))
       
       DTItem = DTItem + dct02
       
       'DESCUENTO ESPECIAL  :w8dct03 =(w8bruto - w8dct02-w8dct06)*w2dctpp/100
        dct03 = (TImporte - dct02 - dct06) * (MBox(7) / 100)            '(Tbruto-dct02-dct06)
        
        DTPPago = DTPPago + dct03
        
       'DESCUENTO POR PROMOCION  : w8dct04 =(w8bruto - w8dct02-w8dct03-w8dct06)*w2dctpr/100
        dct04 = (TImporte - dct02 - dct03 - dct06) * (MBox(6) / 100)
        
        DTPromo = DTPromo + dct04
        
       'DESCUENTO GENERAL : w8dct05 =(w8bruto - w8dct02-w8dct03-w8dct04-w8dct06)*w2dctgl/100
        dct05 = (TImporte - dct02 - dct03 - dct04 - dct06) * (MBox(5) / 100)
        DTGlobal = DTGlobal + dct05
       
       'ACUMULADO DE TOTAL DESCUENTOS  :w8dctos = w8dct02 + w8dct03+w8dct04+w8dct05+w8dct06
        Tdscto = Tdscto + (dct02 + dct03 + dct04 + dct05 + dct06)
        
       'ACUMULADO DE SUBTOTAL DE VENTA : w8subto = w8bruto - w8dctos
        TSub = TSub + (TImporte - (dct02 + dct03 + dct04 + dct05 + dct06))
                
       If parametro.tieneigv = "1" Then
            'CALCULAMOS EL IMPORTE :=  TOTAL IMPORTE SIN IGV - DESCTOS + IGV
            Previo = (TImporte - (dct02 + dct03 + dct04 + dct05 + dct06))
            Previo = (Previo * parametro.igv)
            Tigv = Tigv + Previo
            
            'GRABAMOS EL TOTAL DE IMPORTE EN LA TABLA TEMPORAL PARA MOSTRAR
            Previo = (TImporte - (dct02 + dct03 + dct04 + dct05 + dct06))
            Previo = (Previo * (1 + parametro.igv))
            rsdeta.Fields(7) = Previo
       Else                    'If parametro.tieneigv = "0" Then
           If modoventa.impuestos = "1" Then
                Previo = (TImporte - (dct02 + dct03 + dct04 + dct05 + dct06))
                Previo = (Previo * parametro.igv)
                Tigv = Tigv + Previo
                
                Previo = (TImporte - (dct02 + dct03 + dct04 + dct05 + dct06))
                Previo = (Previo * (1 + parametro.igv))
                rsdeta.Fields(7) = Previo
           Else
               If rsdeta.Fields(11) > 0 Then
                    Previo = (TImporte - (dct02 + dct03 + dct04 + dct05 + dct06))
                    rsdeta.Fields(7) = Previo * (1 + rsdeta(11))
                    Tigv = Tigv + (Previo * rsdeta(11))
               Else
                    Previo = (TImporte - (dct02 + dct03 + dct04 + dct05 + dct06))
                    rsdeta.Fields(7) = Previo
                    'Tigv = Tigv
              End If
           End If
       End If
       rsdeta.Update
       rsdeta.MoveNext
    Next J
  Else
    Exit Function
  End If
  
  ' w2imp = IIf(w2ciaimp, w2timp, pro_pctimp)
  ' w2imp = IIf(vtmod.mod_imp, w2imp, 0)
  ' w2prepac = IIf(w2dctofe > 0, roun(w2prepac * (100 - w2dctofe) / 100, 4), w2prepac
   
   'set deci to 12
   'w8bruto = w2cant  * w2prepac                                          && Total Bruto
   'w2dctofe = 0
   'If w2fchatn>=pro_fchini and w2fchatn<=pro_fchfin                      && Precio de Oferta en una lista de precios
   '   w2dctofe =pro_dctofi      && Descuentos Ofertas
   '*   w8dct01 = w2cant  * Abs(IIF(w2prelis>w2prepac,w2prelis-w2prepac,0))   && Dcto.Oferta
   'w8dct06 = w8bruto * w0dcto/100                                        && Dcto. por Default
   'w8dct02 = (w8bruto-w8dct06)*w2dctlin/100                              && Dcto.Por Item
  
  'IMPORTE TOTAL NETO DE FACTURA   w8tneto = w8subto + w8impto
  TNeto = Tbruto - Tdscto + Tigv
  MBox2(7) = Format(Tbruto, "#,###,##0.0000")
  MBox2(6) = Numero(TCant)
  MBox2(9) = Numero(Tigv)
  MBox2(8) = Numero(Tdscto)
  MBox2(10) = Numero(TNeto)
  Limpiartexto MBox2, 12, 12
  Limpiartexto MBox2, 13, 13
  Limpiartexto MBox2, 14, 14
  Limpiartexto MBox2, 0, 5
  
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
  On Error Resume Next
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
   
    If parametro.tieneigv = "1" Then
         MBox2(3) = Format(TDBGrid1.Columns(5).Text * (1 + (parametro.igv)), "######0.0000")
    Else
       If modoventa.impuestos = "1" Then
           MBox2(3) = Format(IIf(IsNull(TDBGrid1.Columns(5).Text) Or Len(Trim(TDBGrid1.Columns(5).Text)) = 0, 0, TDBGrid1.Columns(5).Text) * (1 + (parametro.igv)), "######0.0000")
       Else
           MBox2(3) = Format(TDBGrid1.Columns(5).Text, "######0.0000")
       End If
    End If
    MBox2(4) = Numero(TDBGrid1.Columns(6).Text)
    MBox2(5) = Numero(TDBGrid1.Columns(8).Text)
    If MBox2(12).Enabled = True Then
      MBox2(12).SetFocus
    Else
      MBox2(0).SetFocus
    End If
    Flag = 1
  End If
  
End Sub

Public Function Carga_Pedido()
    Dim csql As New ADODB.Recordset
    Dim acliente As New ADODB.Recordset
    Dim J As Integer
    Set csql = Cn.Execute("select * from " & g_PedidoPuntoVta & " where pedidonumero='" & TDBGrid2.Columns(0).Text & "'")
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
            MBox(4) = Escadena(csql!pedidonrofact)                   'nro guia
       Else
            MBox(4) = 0
       End If
       MBox(5) = Numero(csql!pedidodsctoglobal)                   'dscto gral
       MBox(6) = Numero(csql!pedidodsctoppago)                    'dscto promocional
       MBox(7) = Numero(csql!pedidodsctovtaoficina)               'dscto especial
       Combo1.ListIndex = VerificaCombo(Combo1, csql!pedidomoneda)     'moneda
       MBox(8) = Numero(csql!pedidotipcambio)                             'tipo de cambio
       Combo2.ListIndex = VerificaCombo(Combo2, Trim(csql!pedidolistaprec)) 'lista precios
       MBox(9) = Escadena(csql!pedidomensaje)                            'mensajes
       Combo3.ListIndex = VerificaCombo(Combo3, csql!modovtacodigo)       'modo de venta
       MBox(10) = Format(csql!pedidofecha, "dd/mm/yyyy")                            'fecha de atencion
       Combo4.ListIndex = VerificaCombo(Combo4, csql!formapagocodigo) 'forma de pago
       Ctr_Ayuda1.xclave = Escadena(csql!clientecodigo)                  ' cliente MBox(11)
       
       '*****Respecto a Clientes *******
       Call Ctr_Ayuda1.Ejecutar
       
       Set acliente = Cn.Execute("select * from vt_cliente where clientecodigo='" & Ctr_Ayuda1.xclave & "'")
       If acliente.RecordCount > 0 Then
          Combo6.ListIndex = VerificaCombo(Combo6, acliente!clientetipopersona)
          Combo7.ListIndex = VerificaCombo(Combo7, acliente!clientetipopais)
          Combo8.ListIndex = VerificaCombo(Combo8, IIf(acliente!clientemultidireccion = 1, "S", "N"))
          lcred(0) = Numero(acliente!clientesaldodolares)
          lcred(1) = Numero(acliente!clientelimitecreddolar)
       End If
       acliente.Close
       Set acliente = Nothing
       
       Ctr_Ayuda2.xclave = Escadena(csql!vendedorcodigo)                    'vendedor
       Call Ctr_Ayuda2.Ejecutar
       MBox(13) = Numero(csql!pedidoporccomision)                           'comision
       Ctr_Ayuda3.xclave = Escadena(csql!almacencodigo)                     'almacen
       Call Ctr_Ayuda3.Ejecutar
       'Ctr_Ayuda3.Filtro = "modovtacodigo in (" & modoventa.almacenes & ")"
       
       MBox(15) = Numero(csql!pedidototalotros)                             'otros gastos
       MBox(16) = Escadena(csql!pedidonotaped)                              'nota pedido
       MBox(17) = Escadena(csql!pedidoordencompra)                          'orden de compra
       Combo5.ListIndex = VerificaCombo(Combo5, IIf(csql!pedidoautorizacion = 1, "S", "N")) 'autorizacion
       MBox(18) = Format(csql!pedidodiaspago, "##0")                        'dias pago
       MBox2(6) = Numero(csql!pedidototitem)                                'Total Cantidad
       MBox2(7) = Numero(csql!pedidototbruto)                               'Total Bruto
       MBox2(8) = Numero(csql!pedidototalflete)                             'Total Dsctos
       MBox2(9) = Numero(csql!pedidototimpuesto)                            'Total Igv
       MBox2(10) = Numero(csql!pedidototneto)                               'Neto a Facturar
       MBox(19) = Escadena(csql!pedidoentrega)                             'Entrega de Pedidos
       Text3 = Escadena(csql!pedidoobserva)
       Text4(0) = Escadena(csql!pedidotiporefe)
       Text4(1) = Escadena(csql!pedidonrorefe)
       
       TClie.Value = 0
       SSTab2.Tab = 0
       SSTab2.TabEnabled(2) = True
    End If
    csql.Close
       
    Set csql = Cn.Execute("select detpeditem,A.productocodigo,b.adescri,a.unidadcodigo," & _
                          "detpedcantpedida,detpedmontoprecvta,detpeddsctoxitem,detpedimpbruto," & _
                          " detpedporccomis,detpedcantpedidaref,detpedfactorconv " & _
                          "from " & g_DetallePuntoVta & " A " & _
                          "inner Join " & _
                          "[" & cbdatos.DefaultDatabase & "].dbo.maeart B " & _
                          " ON A.productocodigo=b.acodigo COLLATE Modern_Spanish_CI_AI " & _
                          "where pedidonumero='" & TDBGrid2.Columns(0).Text & "'")
    
    Set rsdeta = Nothing
    Call CargaGrilla
   
    Do Until csql.EOF
       rsdeta.AddNew
       rsdeta.Fields(0) = Escadena(csql!detpeditem)
       rsdeta.Fields(1) = Escadena(csql!productocodigo)
       rsdeta.Fields(2) = Escadena(csql!adescri)
       rsdeta.Fields(3) = Escadena(csql!unidadcodigo)
       rsdeta.Fields(4) = Numero(csql!detpedcantpedida)
       rsdeta.Fields(5) = Numero(IIf(IsNull(csql!detpedmontoprecvta), 0, csql!detpedmontoprecvta))
       rsdeta.Fields(6) = Numero(csql!detpeddsctoxitem)
       rsdeta.Fields(7) = Numero(csql!detpedimpbruto)
       rsdeta.Fields(8) = Numero(csql!detpedporccomis)
       rsdeta.Fields(9) = Numero(IIf(IsNull(csql!detpedcantpedidaref), 0, csql!detpedcantpedidaref))
       rsdeta.Fields(10) = Numero(IIf(IsNull(csql!detpedfactorconv), 0, csql!detpedfactorconv))
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
    Dim lzona As String
    Dim Previo As Double
    Dim dct02, dct03, dct04, dct05, dct06 As Double
    Dim tinafecto As Double
    Dim xserie As String * 3
    Dim xfactu As Double  'String * 8
    Dim xtipofac As String * 2
    
    Dim acmd As New ADODB.Command
    Dim asql As New ADODB.Recordset
    Dim arbusca As New ADODB.Recordset

    'On Error GoTo vererror
   'On Error Resume Next
    
    Call CargarModo
    
    GrabarData = 0
    
    '******** CABECERA DE MOVIMIENTO *****************
    If rsdeta.RecordCount = 0 Then
      MsgBox W1TXT30, vbInformation, MsgTitle
      GrabarData = 0
      Exit Function
    End If
    Call Totales
    For J = 1 To 29
        wCabe(J) = ""
    Next J
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
    wCabe(12) = MBox(9)                      'mensajes
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
    wCabe(26) = Round(MBox2(7), 2)          'Total Bruto
    wCabe(27) = 0    'MBox2(8)              'total fletes --T.D.
    wCabe(28) = Round(MBox2(9), 2)          'Total Igv
    wCabe(29) = Round(MBox2(10), 2)         'Neto a Facturar
    wCabe(30) = MBox(19)                    'entrega pedido
    wCabe(31) = MBox3(1)                    'nombre cliente
    wCabe(32) = MBox3(3)                    'direccion
    wCabe(33) = MBox3(2)                    'ruc
    wCabe(34) = MBox(10)                     'fechafactura
    wCabe(35) = DTGlobal                     'Total Descuentos Globales
    wCabe(36) = DTCliente                    'Total Descuentos Cliente
    wCabe(37) = DTOficina                    'Total Descuentos Oficina
    wCabe(38) = DTItem                       'Total Descuentos Item
    wCabe(39) = DTLinea                      'Total Descuentos Linea
    wCabe(40) = DTPromo                      'Total Descuentos x Promocion
    wCabe(41) = Trim(Text3)
    wCabe(42) = Trim(Text4(0))
    wCabe(43) = Trim(Text4(1)) & Trim(Text4(2))
    
    If cOpc(0).Value Or cOpc3(0).Value Then
        Set asql = Cn.Execute("select * from " & g_DetallePuntoVta & " where pedidonumero='" & MBox(1) & "'")
        If asql.RecordCount > 0 Then
           Cn.Execute "Delete From " & g_DetallePuntoVta & " where pedidonumero='" & MBox(1) & "'"
           Cn.Execute "Delete From " & g_PedidoPuntoVta & " where pedidonumero='" & MBox(1) & "'"
        End If
        asql.Close
        nsql = "Insert Into " & g_PedidoPuntoVta & "("
    ElseIf cOpc(1).Value Or cOpc3(1).Value Then
        If cOpc3(1).Value Then
           Cn.Execute "Delete From " & g_DetallePuntoVta & " where pedidonumero='" & MBox(1) & "'"
           Cn.Execute "Delete From " & g_PedidoPuntoVta & " where pedidonumero='" & MBox(1) & "'"
        End If
        nsql = "Insert Into vt_Pedido ("
    End If
    Set asql = Nothing
'    cn.CommitTrans
    
'    cn.BeginTrans
    ' ** Verificando Numeracion de Documentos *****
    If cOpc(1).Value Or cOpc3(1).Value Then
        If cOpc2(0).Value Then
          If cOpc(1).Value Then
             MBox(1) = g_pedserie & Right("000000000000" & TraeDataSerie("select puntovtadoccorr from vt_puntovtadocumento where documentocodigo='" & g_tipoped & "' and puntovtadocserie='" & g_pedserie & "' and puntovtacodigo='" & g_ptoventa & "'", Cn), 8) 'MBox(1).MaxLength)
          End If
          Set asql = Cn.Execute("select * from vt_detallepedido where pedidonumero='" & MBox(1) & "'")
          If asql.RecordCount > 0 Then
             Cn.Execute "Delete From vt_detallepedido where pedidonumero='" & MBox(1) & "'"
             Cn.Execute "Delete From vt_pedido where pedidonumero='" & MBox(1) & "'"
          End If
          asql.Close
          'wCabe(34) = Date                       'fechafactura
          MBox(2) = g_facserie & Right("000000000000" & TraeDataSerie("select puntovtadoccorr from vt_puntovtadocumento where documentocodigo='" & g_tipofac & "' and puntovtadocserie='" & g_facserie & "' and puntovtacodigo='" & g_ptoventa & "'", Cn), 8) ' MBox(2).MaxLength)
          MBox(3) = "0": MBox(4) = "0"
          
          If dllgeneral.VerificaDatoExistente(Cn, "select * from vt_pedido where pedidonrofact='" & MBox(2) & "' and pedidotipofac='" & g_tipofac & "'") = 1 Then
            MsgBox "Ya existe Documento " & g_tipofac & "-" & MBox(2), vbInformation, MsgTitle
            GrabarData = 0
            Exit Function
          End If
        ElseIf cOpc2(1).Value Then
          If cOpc(1).Value Then
             MBox(1) = g_pedserie & Right("000000000000" & TraeDataSerie("select puntovtadoccorr from vt_puntovtadocumento where documentocodigo='" & g_tipoped & "' and puntovtadocserie='" & g_pedserie & "' and puntovtacodigo='" & g_ptoventa & "'", Cn), 8) 'MBox(1).MaxLength)
          End If
          'wCabe(34) = Date                       'fechaboleta
          MBox(3) = g_bolserie & Right("000000000000" & TraeDataSerie("select puntovtadoccorr from vt_puntovtadocumento where documentocodigo='" & g_tipobol & "' and puntovtadocserie='" & g_bolserie & "' and puntovtacodigo='" & g_ptoventa & "'", Cn), 8)  'MBox(3).MaxLength)
          MBox(2) = "0": MBox(4) = "0"
          If dllgeneral.VerificaDatoExistente(Cn, "select * from vt_pedido where pedidonrofact='" & MBox(3) & "' and pedidotipofac='" & g_tipobol & "'") = 1 Then
            MsgBox "Ya existe Documento " & g_tipobol & "-" & MBox(3), vbInformation, MsgTitle
            GrabarData = 0
            Exit Function
          End If
        ElseIf cOpc2(2).Value Then
          If cOpc(1).Value Then
             MBox(1) = g_pedserie & Right("000000000000" & TraeDataSerie("select puntovtadoccorr from vt_puntovtadocumento where documentocodigo='" & g_tipoped & "' and puntovtadocserie='" & g_pedserie & "' and puntovtacodigo='" & g_ptoventa & "'", Cn), 8) ' MBox(1).MaxLength)
          End If
         ' wCabe(34) = Date                       'fechaguia
          MBox(4) = g_guiaserie & Right("000000000000" & TraeDataSerie("select puntovtadoccorr from vt_puntovtadocumento where documentocodigo='" & g_tipoguia & "' and puntovtadocserie='" & g_guiaserie & "' and puntovtacodigo='" & g_ptoventa & "'", Cn), 8)   ' MBox(4).MaxLength)
          MBox(2) = "0": MBox(3) = "0"
          If dllgeneral.VerificaDatoExistente(Cn, "select * from vt_pedido where pedidonrofact='" & MBox(3) & "' and pedidotipofac='" & g_tipoguia & "'") = 1 Then
            MsgBox "Ya existe Documento " & g_tipoguia & "-" & MBox(3), vbInformation, MsgTitle
            GrabarData = 0
            Exit Function
          End If
        End If
    End If
    
    If cOpc(1).Value Or cOpc(0).Value Then
        '*** Verifica Serie Documentos *****
        nsql = "Update vt_puntovtadocumento " & _
                " set puntovtadoccorr='" & Right("00000000" & Trim(CStr(MBox(1) + 1)), 8) & "'" & _
                " Where documentocodigo='" & g_tipoped & "' and puntovtacodigo='" & g_ptoventa & "' and puntovtadocserie='" & g_pedserie & "'"
                
        Cn.Execute nsql
    End If
    '***** Actualizando Numeracion de Documentos*****
    If cOpc(1).Value Or cOpc3(1).Value Then
         If cOpc2(0).Value Then
             If Len(Trim(g_facserie)) = 0 Then
                MsgBox "No existe Serie de Facturas....Verifique!!", vbInformation, MsgTitle
                'cn.RollbackTrans
                Exit Function
             End If
            nsql = "Update vt_puntovtadocumento " & _
                  " set puntovtadoccorr='" & Right("00000000" & Trim(CStr(MBox(2) + 1)), 8) & "'" & _
                   " Where documentocodigo='" & g_tipofac & "' and puntovtacodigo='" & g_ptoventa & "' and puntovtadocserie='" & g_facserie & "'"
                   
                   
    
        ElseIf cOpc2(1).Value Then
             If Len(Trim(g_bolserie)) = 0 Then
                MsgBox "No existe Serie de Boletas....Verifique!!", vbInformation, MsgTitle
                'cn.RollbackTrans
                Exit Function
             End If
        
           nsql = "Update vt_puntovtadocumento " & _
                   " set puntovtadoccorr='" & Right("00000000" & Trim(CStr(MBox(3) + 1)), 8) & "'" & _
                   " Where documentocodigo='" & g_tipobol & "' and puntovtacodigo='" & g_ptoventa & "' and puntovtadocserie='" & g_bolserie & "'"
    
        ElseIf cOpc2(2).Value Then
             If Len(Trim(g_guiaserie)) = 0 Then
                MsgBox "No existe Serie de Guias....Verifique!!", vbInformation, MsgTitle
                'cn.RollbackTrans
                Exit Function
             End If
        
             nsql = "Update vt_puntovtadocumento " & _
                    "set puntovtadoccorr='" & Right("00000000" & Trim(CStr(MBox(4) + 1)), 8) & "'" & _
                    " Where documentocodigo='" & g_tipoguia & "' and puntovtacodigo='" & g_ptoventa & "' and puntovtadocserie='" & g_guiaserie & "'"
        End If
        Cn.Execute nsql
        
    End If
    DoEvents
    '**cambio de documentacion
    wCabe(2) = Trim(MBox(1))                         'nro pedido
    wCabe(3) = Trim(MBox(2))                         'nro factura
    wCabe(4) = Trim(MBox(3))                         'nro boleta
    wCabe(5) = Trim(MBox(4))                         'nro guia
    
    If cOpc(1).Value Or cOpc3(1).Value Then
      If cOpc2(0).Value Then
         wCabe(3) = Trim(MBox(2))
         wCabe(4) = g_tipofac
      ElseIf cOpc2(1).Value Then
         wCabe(3) = Trim(MBox(3))
         wCabe(4) = g_tipobol
      ElseIf cOpc2(2).Value Then
         wCabe(3) = Trim(MBox(4))
         wCabe(4) = g_tipoguia
      End If
    Else
        wCabe(3) = 0  ' Trim(MBox(2))                         'nro factura
        wCabe(4) = "" 'Trim(MBox(3))                         'nro boleta
        wCabe(5) = 0  'Trim(MBox(4))                         'nro guia
    End If
    
    DoEvents
    
    Set acmd.ActiveConnection = cg
    acmd.CommandType = adCmdStoredProc
    acmd.CommandText = "vt_ingresapedido_pro"
    acmd.CommandTimeout = 0
    acmd.Prepared = True
    With acmd
        .Parameters("@base") = Cn.DefaultDatabase
        .Parameters("@tabla") = IIf(cOpc(0).Value Or cOpc3(0).Value, g_PedidoPuntoVta, "vt_pedido")
        .Parameters("@tipo") = IIf(dllgeneral.VerificaDatoExistente(Cn, "select * from " & IIf(cOpc(0).Value Or cOpc3(0).Value, g_PedidoPuntoVta, "vt_pedido") & " where pedidonumero='" & wCabe(2) & "'") = 0, "1", "2") '"1"
        .Parameters("@puntovta") = wCabe(1)
        .Parameters("@numero") = wCabe(2)
        .Parameters("@factura") = wCabe(3)
        .Parameters("@boleta") = wCabe(4)
        .Parameters("@guia") = wCabe(5)
        .Parameters("@dsctoglobal") = wCabe(6)
        .Parameters("@dsctoppago") = wCabe(7)
        .Parameters("@dsctovtaofi") = wCabe(8)
        .Parameters("@moneda") = wCabe(9)
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
        .Parameters("@observa") = wCabe(41)
        .Parameters("@tiporefe") = wCabe(42)
        .Parameters("@nrorefe") = wCabe(43)
    End With
    acmd.Execute
    Set acmd = Nothing
    DoEvents
    
    
    
    If modoventa.ctrlinventario = "1" Then
        Set acmd.ActiveConnection = cg
        acmd.CommandType = adCmdStoredProc
        acmd.CommandText = "vt_ingresoalma_pro"
        acmd.CommandTimeout = 0
        acmd.Prepared = True
        With acmd
            .Parameters("@base") = cbdatos.DefaultDatabase
            .Parameters("@tabla") = "movalmcab"
            .Parameters("@tipo") = "1"
            .Parameters("@puntovta") = wCabe(1)
            .Parameters("@numero") = wCabe(2)
            .Parameters("@factura") = wCabe(3)
            .Parameters("@boleta") = wCabe(4)
            .Parameters("@guia") = wCabe(5)
            .Parameters("@dsctoglobal") = wCabe(6)
            .Parameters("@dsctoppago") = wCabe(7)
            .Parameters("@dsctovtaofi") = wCabe(8)
            .Parameters("@moneda") = IIf(wCabe(9) = g_TipoSol, "S", "D")
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
        End With
        acmd.Execute
        Set acmd = Nothing
        DoEvents
    End If
       
    
    If cOpc2(0).Value Or cOpc2(1).Value Or cOpc2(2).Value Then
       If wCabe(9) = g_TipoSol Then
            Cn.Execute "Update vt_cliente " & _
                       " Set clientesaldosoles=ISNULL(clientesaldosoles,0)+" & CDbl(wCabe(29)) & _
                       "      Where clientecodigo='" & wCabe(16) & "'"
       ElseIf wCabe(9) = g_TipoDolar Then
            Cn.Execute "Update vt_cliente " & _
                       " Set clientesaldodolares=ISNULL(clientesaldodolares,0)+" & CDbl(wCabe(29)) & _
                       "      Where clientecodigo='" & wCabe(16) & "'"
       End If
    End If
    DoEvents
    '********** DETALLE DE MOVIMIENTOS *****************
    
    rsdeta.MoveFirst
    Do Until rsdeta.EOF()
       If Val(rsdeta.Fields(0)) = Val(inicio) Then
          Exit Do
        Else
          rsdeta.MoveNext
       End If
    Loop
    regi = 1
    tinafecto = 0
    Do Until rsdeta.EOF
       If Val(rsdeta.Fields(0)) > Val(Fin) Then
          Exit Do
       End If
           'IMPORTE DE MONTO BRUTO SIN IGV, ES DECIR PRECIO X CANTIDAD
           Tbruto = Tbruto + (rsdeta.Fields(4) * rsdeta.Fields(5))
           TCant = TCant + rsdeta.Fields(4)
           TImporte = (rsdeta.Fields(4) * rsdeta.Fields(5))
           
           'DESCUENTO DE CIA O EMPRESA
    '       If parametro.tienedscto = "1" Then
    '            dct06 = TImporte * (1 + (parametro.descuento / 100))
    '       Else
    '          dct06 = 0
    '       End If
           If IsNull(text1) Or Len(Trim(text1)) = 0 Then
                 dct06 = 0
           Else
               'dct06 = TImporte * (1 + parametro.descuento)
               dct06 = TImporte * (CDbl(text1))
           End If
          
           'DESCUENTO POR ITEM
           dct02 = 0
           dct02 = (TImporte * (rsdeta.Fields(6) / 100))
           
           'DESCUENTO ESPECIAL  :w8dct03 =(w8bruto - w8dct02-w8dct06)*w2dctpp/100
            dct03 = 0
            dct03 = (TImporte - dct02 - dct06) * (MBox(7) / 100)            '(Tbruto-dct02-dct06)
            
           'DESCUENTO POR PROMOCION  : w8dct04 =(w8bruto - w8dct02-w8dct03-w8dct06)*w2dctpr/100
            dct04 = 0
            dct04 = (TImporte - dct02 - dct03 - dct06) * (MBox(6) / 100)
            
           'DESCUENTO GENERAL : w8dct05 =(w8bruto - w8dct02-w8dct03-w8dct04-w8dct06)*w2dctgl/100
            dct05 = 0
            dct05 = (TImporte - dct02 - dct03 - dct04 - dct06) * (MBox(5) / 100)
           
           'ACUMULADO DE TOTAL DESCUENTOS  :w8dctos = w8dct02 + w8dct03+w8dct04+w8dct05+w8dct06
            Tdscto = dct02 + dct03 + dct04 + dct05 + dct06
            
           'ACUMULADO DE SUBTOTAL DE VENTA : w8subto = w8bruto - w8dctos
           TSub = 0
           TSub = (TImporte - (dct02 + dct03 + dct04 + dct05 + dct06))
           Previo = TSub
           If parametro.tieneigv = "1" Then
              'CALCULAMOS EL IGV
              Previo = (TSub * parametro.igv)
           Else
                If modoventa.impuestos = "1" Then
                     Previo = (TImporte - (dct02 + dct03 + dct04 + dct05 + dct06))
                     Previo = (Previo * parametro.igv)
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
        
        If cOpc(0).Value Or cOpc3(0).Value Then
            nsql = g_DetallePuntoVta   '"Tempodetallepedido"
        ElseIf cOpc(1).Value Or cOpc3(1).Value Then
            nsql = "vt_detallepedido"
        End If
        Set acmd.ActiveConnection = cg
        acmd.CommandType = adCmdStoredProc
        acmd.CommandTimeout = 0
        acmd.CommandText = "vt_ingresodetallepedido_pro"
        acmd.Prepared = True
        
        With acmd
            .Parameters("@base") = Cn.DefaultDatabase
            .Parameters("@tabla") = nsql
            .Parameters("@tipo") = "1"
            .Parameters("@item") = rsdeta.Fields(0)
            .Parameters("@numero") = MBox(1)
            .Parameters("@producto") = rsdeta.Fields(1)
            .Parameters("@unidad") = rsdeta.Fields(3)
            .Parameters("@cantidad") = rsdeta.Fields(4)
            .Parameters("@preciopacto") = IIf(rsdeta.Fields(6) = 0, rsdeta.Fields(5), (rsdeta.Fields(7) / (100 - rsdeta.Fields(6))) * 100)
            .Parameters("@dsctoxitem") = rsdeta.Fields(6)
            .Parameters("@importebruto") = rsdeta.Fields(7)
            .Parameters("@porcomision") = rsdeta.Fields(8)
            .Parameters("@mdsctoitem") = Tdscto
            .Parameters("@mdsctoxlinea") = 0
            .Parameters("@mdsctoxprom") = Previo     '0
            .Parameters("@mimpor") = rsdeta.Fields(7)       'Previo
            .Parameters("@unidadref") = IIf(IsNull(rsdeta.Fields(9)) Or Len(Trim(rsdeta.Fields(9))) = 0, 0, CDbl(rsdeta.Fields(9)))
            .Parameters("@preciolista") = Val(IIf(IsNull(rsdeta.Fields(12)), 0, IIf(Len(Trim(rsdeta.Fields(12))) = 0, 0, rsdeta.Fields(12))))
            .Parameters("@partida") = " "
            .Parameters("@metrica") = " "
       
        End With
        acmd.Execute
        Set acmd = Nothing
            
            '******Actualizamos Saldos en Almacen *********
            If modoventa.ctrlinventario = "1" Then
            
                '--Actualizamos el archivo facart--
               If cOpc2(0).Value Or cOpc2(1).Value Or cOpc2(2).Value Then
               
                    If cOpc2(0).Value Then
                        xserie = Left(MBox(2).Text, 3)
                        xfactu = Val(Right(MBox(2).Text, 8))
                        xtipofac = g_tipofac
                    ElseIf cOpc2(1).Value Then
                        xserie = Left(MBox(3).Text, 3)
                        xfactu = Val(Right(MBox(3).Text, 8))
                        xtipofac = g_tipobol
                    ElseIf cOpc2(2).Value Then
                        xserie = Left(MBox(4).Text, 3)
                        xfactu = Val(Right(MBox(4).Text, 8))
                        xtipofac = g_tipoguia
                    Else
                        xserie = Left(MBox(1).Text, 3)
                        xfactu = Val(Right(MBox(1).Text, 8))
                        xtipofac = g_tipoped
                    End If
                    
                    Set acmd.ActiveConnection = cg
                    acmd.CommandType = adCmdStoredProc
                    acmd.CommandTimeout = 0
                    acmd.CommandText = "vt_ingresodetallealma_pro"
                    acmd.Prepared = True
                    With acmd
                        .Parameters("@base") = cbdatos.DefaultDatabase
                        .Parameters("@tabla") = "movalmdet" ' nsql
                        .Parameters("@tipo") = "1"
                        .Parameters("@item") = rsdeta.Fields(0)
                        .Parameters("@numero") = MBox(1)
                        .Parameters("@producto") = Trim(rsdeta.Fields(1))
                        .Parameters("@unidad") = rsdeta.Fields(3)
                        .Parameters("@cantidad") = rsdeta.Fields(4)
                        .Parameters("@preciopacto") = rsdeta.Fields(5)
                        .Parameters("@dsctoxitem") = rsdeta.Fields(6)
                        .Parameters("@importebruto") = rsdeta.Fields(7)
                        .Parameters("@porcomision") = rsdeta.Fields(8)
                        .Parameters("@mdsctoitem") = Tdscto
                        .Parameters("@mdsctoxlinea") = 0
                        .Parameters("@mdsctoxprom") = Previo     '0
                        .Parameters("@mimpor") = rsdeta.Fields(7)       'Previo
                        .Parameters("@unidadref") = IIf(IsNull(rsdeta.Fields(9)) Or Len(Trim(rsdeta.Fields(9))) = 0, 0, CDbl(rsdeta.Fields(9)))
                        .Parameters("@almacen") = wCabe(19)
                    End With
                    acmd.Execute
                    Set acmd = Nothing
               
                    Set acmd.ActiveConnection = cg
                    acmd.CommandType = adCmdStoredProc
                    acmd.CommandTimeout = 0
                    acmd.CommandText = "vt_actualizoalma_pro"
                    acmd.Prepared = True
                    With acmd
                        .Parameters("@basedes") = CStr(cbdatos.DefaultDatabase)
                        .Parameters("@almacen") = wCabe(19)
                        .Parameters("@tipo") = "1"
                        .Parameters("@articulo") = Trim(rsdeta.Fields(1))
                        .Parameters("@cantidad") = rsdeta.Fields(4)
                    End With
                    acmd.Execute
                    Set acmd = Nothing
             End If
        End If
                
        rsdeta.MoveNext
        regi = regi + 1
    Loop
    
    '*****Actualizamos el Valor de Inafecto**********
    Cn.Execute "UPDATE " & g_PedidoPuntoVta & _
               " Set Pedidototinafecto=" & tinafecto & _
               " Where pedidonumero='" & MBox(1) & "'"
    
   '*Grabar en los cargos ***ctacte ***
    
    If (cOpc2(0).Value Or cOpc2(1).Value Or cOpc2(2).Value) And modoventa.ctacte = "1" Then
        lzona = "00"
        Set asql = Cn.Execute("select * from vt_zonavendedor where vendedorcodigo='" & wCabe(17) & "'")
        If asql.RecordCount > 0 Then
            lzona = Escadena(asql!zonacodigo)
        End If
        asql.Close
        Set asql = Nothing
           
        ltipo = "1"
        If dllgeneral.VerificaDatoExistente(Cn, "select * from vt_cargo where documentocargo='" & IIf(cOpc2(0).Value, g_tipofac, g_tipobol) & "' and cargonumdoc='" & IIf(cOpc2(0).Value, MBox(3), MBox(4)) & "'") = 0 Then
          ltipo = "1"
        Else
          ltipo = "2"
        End If
        Set acmd.ActiveConnection = cg
        acmd.CommandType = adCmdStoredProc
        acmd.CommandTimeout = 0
        acmd.CommandText = "vt_ingresacargofactura_pro"
        acmd.Prepared = True
        With acmd
            .Parameters("@base") = Cn.DefaultDatabase
            .Parameters("@tipo") = ltipo
            .Parameters("@tabla") = "vt_cargo"
            If cOpc2(0).Value = True Then
                .Parameters("@tipodocu") = g_tipofac
                .Parameters("@numero") = MBox(2)
            ElseIf cOpc2(1).Value = True Then
                .Parameters("@tipodocu") = g_tipobol
                .Parameters("@numero") = MBox(3)
            Else
                .Parameters("@tipodocu") = g_tipoguia
                .Parameters("@numero") = MBox(4)
            End If
            .Parameters("@cliente") = Escadena(wCabe(16))
            .Parameters("@vendedor") = Escadena(wCabe(17))
            .Parameters("@zona") = lzona
            .Parameters("@apefecemi") = wCabe(14)
            .Parameters("@moneda") = Escadena(wCabe(9))
            .Parameters("@apeimppag") = wCabe(29)
            .Parameters("@usuario") = g_usuario
            .Parameters("@tipocambio") = wCabe(10)
            .Parameters("@fechaact") = Date
            .Parameters("@flagcancel") = "0"
            .Parameters("@fechavenci") = CDate(wCabe(14)) + CDbl(wCabe(24))
            .Parameters("@cargoabono") = "C"
        End With
        acmd.Execute
        Set acmd = Nothing
        
    End If
    If cOpc(1).Value Or cOpc3(1).Value Then
         If cOpc2(0).Value Then
              MsgBox "Se Grabo Satisfactoriamente el Documento." & Chr(13) & Chr(10) & "FACTURA => " & MBox(2), vbInformation, MsgTitle
         ElseIf cOpc2(1).Value Then
              MsgBox "Se Grabo Satisfactoriamente el Documento." & Chr(13) & Chr(10) & "BOLETA => " & MBox(3), vbInformation, MsgTitle
         ElseIf cOpc2(2).Value Then
              MsgBox "Se Grabo Satisfactoriamente el Documento." & Chr(13) & Chr(10) & "GUIA => " & MBox(4), vbInformation, MsgTitle
         Else
              MsgBox "Se Grabo Satisfactoriamente el Documento." & Chr(13) & Chr(10) & "PEDIDO => " & MBox(1), vbInformation, MsgTitle
         End If
     End If
    GrabarData = 1
    
    
'vererror:
'   If Err Then
'      MsgBox "Comunicarse con Sistemas ...!!" & Chr(13) & Chr(10) & cn.Errors(0).Number & "-" & cn.Errors(0).Description
'      Exit Function
'   End If
End Function






Public Function verificaproducto() As Integer
   On Error Resume Next
    verificaproducto = 0
    If rsdeta.RecordCount > 0 Then
       rsdeta.MoveFirst
       Do Until rsdeta.EOF
           If Escadena(rsdeta.Fields(1)) = MBox2(1) And Flag = 0 Then
              verificaproducto = 1
              Exit Do
           End If
           rsdeta.MoveNext
       Loop
    End If
End Function

Public Sub TraerProducto()
  Dim rabusca As New ADODB.Recordset
  Dim nsql As String
  Dim nvalor As Double
  Dim mone As String
  Dim nprecio As Double
  On Error Resume Next

    If Combo2.ListCount > 0 Then
       If modoventa.ctrlinventario = 0 Then
            nsql = "select *,stskdis from [" & cbdatos.DefaultDatabase & "].dbo.maeart " & _
                    "inner join [" & _
                    cbdatos.DefaultDatabase & "].dbo.stkart " & _
                    " ON acodigo=stcodigo " & _
                   " where acodigo='" & MBox2(1) & "'"
       Else
            nsql = "select *,stskdis from [" & cbdatos.DefaultDatabase & "].dbo.maeart " & _
                    "inner join [" & _
                    cbdatos.DefaultDatabase & "].dbo.stkart " & _
                    " ON acodigo=stcodigo " & _
                   " where acodigo='" & MBox2(1) & "' and stalma='" & Ctr_Ayuda3.xclave & "'"
       End If
    Else
       If modoventa.ctrlinventario = 0 Then
            nsql = "select *,stskdis from [" & cbdatos.DefaultDatabase & "].dbo.maeart " & _
                    "inner join [" & _
                    cbdatos.DefaultDatabase & "].dbo.stkart " & _
                    " ON acodigo=stcodigo " & _
                   " where acodigo='" & MBox2(1) & "'"
       Else
            nsql = "select *,stskdis from [" & cbdatos.DefaultDatabase & "].dbo.maeart " & _
                    "inner join [" & _
                    cbdatos.DefaultDatabase & "].dbo.stkart " & _
                    " ON acodigo=stcodigo " & _
                   " where acodigo='" & MBox2(1) & "' and stalma='" & Ctr_Ayuda3.xclave & "'"
        End If
    End If
    Set rabusca = Cn.Execute(nsql)
    If rabusca.RecordCount > 0 Then
      If Val(rabusca.Fields("stskdis")) < Val(MBox2(0).ClipText) Then
        If modoventa.ctrlinventario = "1" Then
            nvalor = Abs(IIf(Val(rabusca.Fields("stskdis")) = 0, 1, Val(rabusca.Fields("stskdis"))))
            If Trim(MBox2(1).ClipText) = "000" Then
                MBox2(1) = Trim(MBox2(1))
            End If
            If (Abs(Val(rabusca.Fields("stskdis")) - Val(MBox2(0).ClipText)) / nvalor) > 0.00025 And Trim(MBox2(1).ClipText) <> "000" Then
              MsgBox "La cantidad disponible es ==>" & Numero(rabusca.Fields("stskdis")) & "...Verifique!!!", vbInformation, "AVISO"
              rabusca.Close
              Set rabusca = Nothing
              Exit Sub
            End If
         End If
      End If
      Label2 = Escadena(rabusca!adescri)
      MBox2(2) = Escadena(rabusca!aunidad)
      If rabusca!acodmon = "MN" Then
         mone = g_TipoSol
      ElseIf rabusca!acodmon = "ME" Then
         mone = g_TipoDolar
      Else
         mone = rabusca!acodmon
      End If
      If mone <> dllgeneral.ComboDato(Combo1.Text) Then
         If dllgeneral.ComboDato(Combo1.Text) = g_TipoSol Then
            nprecio = TraePrecio(Combo2.Text, MBox2(1).Text, Cn, Trim(Ctr_Ayuda3.xclave))
            If nprecio > 0 Then
               MBox2(3) = Numero(nprecio * CDbl(MBox(8)))
               If modoventa.unidadmedida = "R" And modoventa.usafactor = "1" Then
                  MBox2(0) = Numero(0)  'rabusca!unidadfactorconv)
                  MBox2(13) = Numero(0) 'rabusca!unidadfactorconv)
               ElseIf modoventa.unidadmedida = "V" And modoventa.usafactor = "1" Then
                  MBox2(13) = Numero(0)  'rabusca!unidadfactorconv)
               Else
                  MBox2(13) = 1
               End If
               MBox2(14) = Numero(0) 'rabusca!productoporcimpto)
            Else
               MBox2(3) = Numero(TraePrecio(Combo2.Text, MBox2(1).Text, Cn, Trim(Ctr_Ayuda3.xclave)))  'rabusca!productoprecvta)
               If modoventa.unidadmedida = "R" And modoventa.usafactor = "1" Then
                  MBox2(0) = Numero(0)  'rabusca!unidadfactorconv)
                  MBox2(13) = Numero(0) 'rabusca!unidadfactorconv)
               ElseIf modoventa.unidadmedida = "V" And modoventa.usafactor = "1" Then
                  MBox2(13) = Numero(0)  'rabusca!unidadfactorconv)
               Else
                  MBox2(13) = 1
               End If
               MBox2(14) = Numero(0) 'rabusca!productoporcimpto)
            End If
         ElseIf dllgeneral.ComboDato(Combo1.Text) = g_TipoDolar Then
            nprecio = TraePrecio(Combo2.Text, MBox2(1).Text, Cn, Trim(Ctr_Ayuda3.xclave))
            If nprecio > 0 Then
               MBox2(3) = Numero(nprecio / CDbl(MBox(8)))
               If modoventa.unidadmedida = "R" And modoventa.usafactor = "1" Then
                  MBox2(0) = Numero(0)   'rabusca!unidadfactorconv)
                  MBox2(13) = Numero(0)  'rabusca!unidadfactorconv)
               ElseIf modoventa.unidadmedida = "V" And modoventa.usafactor = "1" Then
                  MBox2(13) = Numero(0)  'rabusca!unidadfactorconv)
               Else
                  MBox2(13) = 1
               End If
               MBox2(14) = Numero(0)     'rabusca!productoporcimpto)
            Else
               MBox2(3) = Numero(TraePrecio(Combo2.Text, MBox2(1).Text, Cn, Trim(Ctr_Ayuda3.xclave))) 'rabusca!productoprecvta)
               If modoventa.unidadmedida = "R" And modoventa.usafactor = "1" Then
                  MBox2(0) = Numero(0)   'rabusca!unidadfactorconv)
                  MBox2(13) = Numero(0)  'rabusca!unidadfactorconv)
               ElseIf modoventa.unidadmedida = "V" And modoventa.usafactor = "1" Then
                  MBox2(13) = Numero(0)  'rabusca!unidadfactorconv)
               Else
                  MBox2(13) = 1
               End If
               MBox2(14) = Numero(0)   'rabusca!productoporcimpto)
            End If
         End If
      Else
         MBox2(3).Text = Numero(TraePrecio(Combo2.Text, MBox2(1).Text, Cn, Trim(Ctr_Ayuda3.xclave))) 'rabusca!productoprecvta)
         MBox2(3).Tag = Numero(TraePrecio(Combo2.Text, MBox2(1).Text, Cn, Trim(Ctr_Ayuda3.xclave))) 'rabusca!productoprecvta)
         If modoventa.unidadmedida = "R" And modoventa.usafactor = "1" Then
            MBox2(0) = Numero(0)    'rabusca!unidadfactorconv)
            MBox2(13) = Numero(0)   'rabusca!unidadfactorconv)
         ElseIf modoventa.unidadmedida = "V" And modoventa.usafactor = "1" Then
            MBox2(13) = Numero(0)   'rabusca!unidadfactorconv)
         Else
            MBox2(13) = 1
         End If
         MBox2(14) = Numero(0)   'rabusca!productoporcimpto)
         
      End If
    Else
      Label2 = "":    MBox2(2) = ""
    End If
    MBox2(4) = Numero(0)
    MBox2(5) = Numero(0)
    rabusca.Close
    Set rabusca = Nothing
End Sub

Public Function DocImprimir()
   Dim rf As New ADODB.Command
   Dim rb As New ADODB.Recordset
   Dim puntero, nguia As String
   Dim cuenta As Double
   Dim ntabla As String
   Dim J As Integer
   Dim busca As New dll_apisgen.dll_apis
   
   'On Error Resume Next
      
  Cn.Execute "delete from gtempfile"
  Cn.Execute "delete from tempfile"
  If cOpc2(0).Value Then
     ntabla = "vt_detallepedido"
   Else
     If cOpc2(1).Value Then
        ntabla = "vt_detallepedido"
      Else
        If cOpc2(2).Value Then
           ntabla = "vt_detallepedido"
         Else
            ntabla = g_DetallePuntoVta
        End If
     End If
   End If
  
  Cn.Execute "INSERT into gtempfile" & _
         " Select a.detpedcantpedida,a.productocodigo,b.adescri,(a.detpedimpbruto/a.detpedcantpedida),a.detpedimpbruto,a.detpeddsctoxitem,isnull(a.detpedcantpedidaref,0), case ltrim(rtrim(a.productocodigo)) when '000' then '' else a.unidadcodigo end" & _
         " From " & ntabla & " A inner join " & _
         "[" & cbdatos.DefaultDatabase & "].dbo.maeart B" & _
         " ON A.productocodigo=b.acodigo COLLATE Modern_Spanish_CI_AI " & _
         " Where pedidonumero='" & CStr(MBox(1)) & "'"
      
        
   Call impresion_pedido

   nguia = "000000000"
   If modoventa.emiteguia = "1" Then
        Call procImprimirguia
  ' cn.Execute "Delete from tempfile"
   
'   cn.Execute "drop table tempfile"
'   cn.Execute "Create table tempfile" & _
'             "( detpedcantpedida char(8)," & _
'             " productocodigo char(8)," & _
'             " productodescripcion char(80)," & _
'             " detpedmontoprecvta float," & _
'             "  detpedimpbruto float," & _
'             "  detpeddsctoxitem float," & _
'             "  detpedfactorconv float," & _
'             "  unidadcodigo char(3))"
    End If
    If cOpc2(0).Value Then
       ntabla = "vt_detallepedido"
     Else
       If cOpc2(1).Value Then
          ntabla = "vt_detallepedido"
        Else
          If cOpc2(2).Value Then
             ntabla = "vt_detallepedido"
          Else
             ntabla = g_DetallePuntoVta
          End If
       End If
    End If
    
    Cn.Execute "INSERT into tempfile" & _
             " Select a.detpedcantpedida,a.productocodigo,b.adescri, case when a.detpeddsctoxitem=0 then  (a.detpedimpbruto/a.detpedcantpedida) else detpedmontoprecvta end,a.detpedimpbruto,a.detpeddsctoxitem,isnull(a.detpedcantpedidaref,0),case ltrim(rtrim(a.productocodigo)) when '000' then '' else B.Aunidad end " & _
             " From " & ntabla & " A inner join " & _
            "[" & cbdatos.DefaultDatabase & "].dbo.maeart B " & _
            " ON A.productocodigo=b.acodigo COLLATE Modern_Spanish_CI_AI " & _
            " Where pedidonumero='" & CStr(MBox(1)) & "'"

'    Set rb = Cn.Execute("select * from tempfile")
'    If rb.RecordCount > 0 Then
'     cuenta = modoventa.nroitem - rb.RecordCount
'     If cuenta >= 1 Then
'        For J = 1 To cuenta
'           Cn.Execute "INSERT INTO tempfile(detpedcantpedida,productocodigo,productodescripcion,detpedmontoprecvta,detpedimpbruto,detpeddsctoxitem,detpedfactorconv,unidadcodigo)" & _
'                      " VALUES ( ' ',' ',' ',null,null,null,null,' ')"
'        Next J
'     End If
'   Else
'       Exit Sub
'   End If
'   rb.Close
'   Set rb = Nothing
   
   
   oCrystalReport.Reset
   If cOpc2(0).Value Then
      oCrystalReport.ReportFileName = RutaRep & "Repfactuimpresa.rpt"
    Else
         If cOpc2(1).Value Then
           oCrystalReport.ReportFileName = RutaRep & "Repboletaimpresa.rpt"
         Else
             If cOpc2(2).Value Then
               oCrystalReport.ReportFileName = RutaRep & "Repboimpresa.rpt"
             Else
               oCrystalReport.ReportFileName = RutaRep & "Reppedido.rpt"
             End If
         End If
   End If
   oCrystalReport.LogOnServer "pdssql.dll", _
         busca.LeerIni(App.Path & "\Camtex.ini", "Bventas", "dserver", ""), _
         busca.LeerIni(App.Path & "\Camtex.ini", "Bventas", "dbase", ""), _
         busca.LeerIni(App.Path & "\Camtex.ini", "Bventas", "duser", ""), _
         busca.LeerIni(App.Path & "\Camtex.ini", "Bventas", "dpass", "")
   oCrystalReport.Connect = _
        "DSN=" & busca.LeerIni(App.Path & "\Camtex.ini", "Bventas", "dserver", "") & ";" & _
        "DSQ=" & busca.LeerIni(App.Path & "\Camtex.ini", "Bventas", "dbase", "") & ";" & _
        "UID=" & busca.LeerIni(App.Path & "\Camtex.ini", "Bventas", "duser", "") & ";" & _
        "PWD=" & busca.LeerIni(App.Path & "\Camtex.ini", "Bventas", "dpass", "")
           
   oCrystalReport.Destination = crptToWindow
   oCrystalReport.WindowState = crptMaximized
   oCrystalReport.DiscardSavedData = True
   With oCrystalReport
       If cOpc2(0).Value Then
          .Formulas(0) = "nro='" & MBox(2) & "'"
          .Formulas(8) = "cpago='" & Escadena(Combo4) & "'"
          .Formulas(9) = "opedido='" & MBox(1) & "'"
          .Formulas(10) = "ocompra='" & MBox(17) & "'"
          .Formulas(11) = "guia='" & nguia & "'"
          .Formulas(12) = "fechaven='" & CStr(CDate(MBox(10)) + CDbl(MBox(18))) & "'"
          .Formulas(13) = "vendedor='" & Escadena(Ctr_Ayuda2.xnombre) & "'"
          .Formulas(14) = "bruto='" & Numero(MBox2(7)) & "'"
          .Formulas(15) = "dscto='" & MBox2(8) & "'"
          .Formulas(16) = "igv='" & MBox2(9) & "'"
          .Formulas(17) = "distrito='" & MBox3(4) & "'"
          .Formulas(18) = "observa='" & Trim(Text3) & "'"
       ElseIf cOpc2(1).Value Then
          .Formulas(0) = "nro='" & MBox(3) & "'"
       ElseIf cOpc2(2).Value Then
          .Formulas(0) = "nro='" & MBox(4) & "'"
       Else
          .Formulas(0) = "nro='" & MBox(1) & "'"
       End If
       .Formulas(1) = "cliente='" & Trim(MBox3(1)) & "'"
       .Formulas(2) = "fecha='" & CStr(MBox(10)) & "'"
       .Formulas(3) = "direccion='" & "" & Trim(MBox3(3)) & "'"
       .Formulas(4) = "dni='" & "" & Trim(MBox3(2)) & "'"
       If cOpc2(0).Value Or cOpc2(1).Value Or cOpc2(2).Value Then
          If cOpc2(0).Value Then
            .Formulas(5) = "letras= '" & dllgeneral.NUMLET(Numero(Round(CDbl(MBox2(10)), 2))) & IIf(dllgeneral.ComboDato(Combo1.Text) = g_TipoSol, "Nuevos Soles", "Dolares Americanos") & "'"
          Else
            .Formulas(5) = "letras= '" & "SON : " & dllgeneral.NUMLET(Numero(Round(CDbl(MBox2(10)), 2))) & IIf(dllgeneral.ComboDato(Combo1.Text) = g_TipoSol, "Nuevos Soles", "Dolares Americanos") & "'"
          End If
       Else
         .Formulas(5) = "letras= '" & "SON : " & dllgeneral.NUMLET(Round(CDbl(MBox2(10)), 2)) & IIf(dllgeneral.ComboDato(Combo1.Text) = g_TipoSol, "Nuevos Soles", "Dolares Americanos") & "'"
         .Formulas(6) = "forma='" & Combo4.Text & "'"
         .Formulas(7) = "moneda='" & Combo1.Text & "'"
       End If
   End With
   oCrystalReport.Action = 1
End Function
Private Sub impresion_pedido()
Dim contador As Double
Dim rb As New ADODB.Recordset
Dim busca As New dll_apisgen.dll_apis
Dim nguia As String
Dim SQL As String
nguia = "xx"
Cn.Execute "delete from gtempfile2filas"
Set rb = Cn.Execute("select * from gtempfile")
If rb.RecordCount > 0 Then
   rb.MoveFirst
   contador = 0
   Do Until rb.EOF
      contador = contador + 1
      If contador <= 50 Then
         SQL = "INSERT INTO gtempfile2filas(item,producto1,descripcion1,cantidad1,importe1)"
         SQL = SQL & " VALUES ( '" & contador & "','" & RTrim(rb!productocodigo) & "','" & RTrim(rb!productodescripcion) & "','" & rb!detpedcantpedida & "','" & rb!detpedimpbruto & "')"
         
       Else
        TCant = contador - 50
         SQL = "UPDATE gtempfile2filas set producto2 ='" & RTrim(rb!productocodigo) & "',"
         SQL = SQL & " descripcion2='" & RTrim(rb!productodescripcion) & "',"
         SQL = SQL & "cantidad2='" & rb!detpedcantpedida & "',"
         SQL = SQL & "importe2= '" & rb!detpedimpbruto & "'"
         SQL = SQL & " where item = " & TCant & ""
      End If
      Cn.Execute SQL
      rb.MoveNext
   Loop
   rb.Close
   Set rb = Nothing
End If

oCrystalReport.Reset
oCrystalReport.ReportFileName = RutaRep & "vt_pedido.rpt"
oCrystalReport.LogOnServer "pdssql.dll", _
 busca.LeerIni(App.Path & "\Camtex.ini", "Bventas", "dserver", ""), _
 busca.LeerIni(App.Path & "\Camtex.ini", "Bventas", "dbase", ""), _
 busca.LeerIni(App.Path & "\Camtex.ini", "Bventas", "duser", ""), _
 busca.LeerIni(App.Path & "\Camtex.ini", "Bventas", "dpass", "")
oCrystalReport.Connect = _
 "DSN=" & busca.LeerIni(App.Path & "\Camtex.ini", "Bventas", "dserver", "") & ";" & _
 "DSQ=" & busca.LeerIni(App.Path & "\Camtex.ini", "Bventas", "dbase", "") & ";" & _
 "UID=" & busca.LeerIni(App.Path & "\Camtex.ini", "Bventas", "duser", "") & ";" & _
 "PWD=" & busca.LeerIni(App.Path & "\Camtex.ini", "Bventas", "dpass", "")
                
 oCrystalReport.Destination = crptToWindow
 oCrystalReport.WindowState = crptMaximized
 oCrystalReport.DiscardSavedData = True
 With oCrystalReport
      .Formulas(0) = "nro='" & MBox(2) & "'"
      .Formulas(1) = "cliente='" & MBox3(1) & "'"
      .Formulas(2) = "fecha='" & CDate(MBox(10)) & "' "
      .Formulas(3) = "direccion='" & MBox3(3) & "'"
      .Formulas(4) = "dni='" & MBox3(2) & "'"
      .Formulas(5) = "opedido='" & MBox(1) & "'"
      .Formulas(6) = "ocompra='" & MBox(17) & "'"
      .Formulas(7) = "guia='" & nguia & "'"
      .Formulas(8) = "distrito='" & MBox3(4).ClipText & "'"
      .Formulas(9) = "destino='" & MBox(19).ClipText & "'"
      Set rb = Cn.Execute("select * from gr_empresa where empresacodigo='" & g_Empresa & "'")
      If rb.RecordCount > 0 Then
        .Formulas(10) = "partida='" & Escadena(rb!empresadireccion) & "'"
       Else
         .Formulas(10) = "partida=''"
      End If
      rb.Close
      Set rb = Nothing
 End With
 oCrystalReport.Action = 1
 
End Sub
  

Public Sub CargarModo()
     Dim rs As New ADODB.Recordset
     Dim ncade As String
     Dim J As Integer

     On Error Resume Next
     Set rs = Cn.Execute("select * from vt_modoventa where modovtacodigo='" & dllgeneral.ComboDato(Combo3.Text) & "'")
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
        modoventa.unidadmedida = Escadena(IIf(IsNull(rs!modovtaunidadmedida) Or rs!modovtaunidadmedida = "V", "V", Escadena(rs!modovtaunidadmedida)))
        modoventa.unidadmedida = Left(modoventa.unidadmedida, 1)
        modoventa.usafactor = Escadena(IIf(IsNull(rs!modovtausafactconv) Or rs!modovtausafactconv = 0, "0", "1"))
        If Not IsNull(rs!modovtaalmacen) Then
           ncade = "'"
           For J = 1 To Len(Trim(rs!modovtaalmacen))
             If Mid(Trim(rs!modovtaalmacen), J, 1) <> "," Then
                 ncade = ncade & Mid(Trim(rs!modovtaalmacen), J, 1)
             Else
                 ncade = ncade & "','"
             End If
           Next J
           ncade = ncade & "'"
           modoventa.almacenes = ncade
        Else
           modoventa.almacenes = ""
        End If
        
        MBox(1).Enabled = IIf(modoventa.documento = g_tipoped And modoventa.numeraauto <> "1" And modoventa.ingpedido = "1", True, False) 'Modo de pedido
        MBox(2).Enabled = IIf(modoventa.documento = g_tipofac And modoventa.numeraauto <> "1", True, False) 'Modo de factura
        MBox(3).Enabled = IIf(modoventa.documento = g_tipobol And modoventa.numeraauto <> "1", True, False) 'Modo de boleta
        MBox(4).Enabled = IIf(modoventa.documento = g_tipoguia And modoventa.numeraauto <> "1" And modoventa.ingguia = "1", True, False)  'Modo de Modifica
        
        modoventa.numeraauto = Escadena(IIf(IsNull(rs!modovtanumautom) Or rs!modovtanumautom = 0, "0", "1"))
        modoventa.documento = Escadena(IIf(IsNull(rs!documentocodigo), "", rs!documentocodigo))
        
        MBox2(0).Enabled = IIf(modoventa.usafactor = 0 Or (modoventa.usafactor = "1" And modoventa.unidadmedida = "V"), True, False)
        MBox2(12).Enabled = IIf(modoventa.usafactor = 0 Or (modoventa.usafactor = "1" And modoventa.unidadmedida = "R"), True, False)
     End If
     rs.Close
     Set rs = Nothing

End Sub
Private Sub TDBGrid3_Click()
   If rsmasivo.RecordCount > 0 Then
      TDBGrid3.SetFocus
   End If
End Sub

Private Sub TDBGrid3_KeyDown(KeyCode As Integer, Shift As Integer)
  Dim nvalor As Variant
  On Error Resume Next
  nvalor = KeyCode
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Or KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
   Text10 = Text10 - Val(rsmasivo!cantidad) + Val(TDBGrid3.Columns(5))
  If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
    rsmasivo.MoveNext
  End If
End If
End Sub


Private Sub Text3_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     SendKeys "{tab}"
  End If
End Sub

Private Sub Text4_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Index Like "[12]" Then
            Text4(Index) = Right("0000000000" & Trim(Text4(Index)), Text4(Index).MaxLength)
            If Index = 2 And Len(Trim(Text4(0))) > 0 Then
                If dllgeneral.VerificaDatoExistente(Cn, "select * from vt_pedido where pedidotipofac='" & Text4(0) & "' and pedidonrofact='" & Trim(Text4(1)) & Trim(Text4(2)) & "'") = 0 Then
                    MsgBox "No existe documento...Verifique!!!", vbInformation, "AVISO"
                    Exit Sub
                End If
            ElseIf Chkmasivo = 1 Then
              Call loadmasivo
              
              Call dllgeneral.ActivaTab(2, 2, SSTab1)
              TDBGrid3.SetFocus
            Else
              MBox2(0).SetFocus
            End If
        End If
        If Chkmasivo = 0 Then
            SendKeys "{tab}"
        ElseIf Index Like "[01]" Then
            SendKeys "{tab}"
        End If
    End If

End Sub
