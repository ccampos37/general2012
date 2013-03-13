VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{4D137D9C-00A6-4458-9B46-7E95DB76D55B}#9.0#0"; "TextFer.ocx"
Begin VB.Form PrcCorrelativos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Procesos de Correlativos"
   ClientHeight    =   9600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9405
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9600
   ScaleWidth      =   9405
   Begin TabDlg.SSTab SSTab1 
      Height          =   8880
      Left            =   240
      TabIndex        =   5
      Top             =   210
      Width           =   8925
      _ExtentX        =   15743
      _ExtentY        =   15663
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      ForeColor       =   8388608
      TabCaption(0)   =   "Documento Venta"
      TabPicture(0)   =   "PrcCorrelativos.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame4(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Fr1(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Fr2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame2(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Documento Anulado"
      TabPicture(1)   =   "PrcCorrelativos.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2(1)"
      Tab(1).Control(1)=   "Frame4(1)"
      Tab(1).Control(2)=   "Fr1(1)"
      Tab(1).Control(3)=   "Frame1"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Ingresa No. Pedido"
      TabPicture(2)   =   "PrcCorrelativos.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame2(2)"
      Tab(2).Control(1)=   "Frame4(2)"
      Tab(2).Control(2)=   "Fr1(2)"
      Tab(2).Control(3)=   "Frame3"
      Tab(2).ControlCount=   4
      Begin VB.Frame Frame2 
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
         Height          =   915
         Index           =   2
         Left            =   -74550
         TabIndex        =   120
         Top             =   6300
         Width           =   2775
         Begin VB.Label Label31 
            Caption         =   "Ayuda de Doc."
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
            Left            =   1380
            TabIndex        =   122
            Top             =   390
            Width           =   1185
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   2
            Left            =   150
            Picture         =   "PrcCorrelativos.frx":0054
            Top             =   330
            Width           =   480
         End
         Begin VB.Label Label30 
            Caption         =   "[ F1 ]"
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
            Index           =   3
            Left            =   900
            TabIndex        =   121
            Top             =   390
            Width           =   435
         End
      End
      Begin VB.Frame Frame2 
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
         Height          =   915
         Index           =   1
         Left            =   -74580
         TabIndex        =   117
         Top             =   6330
         Width           =   2835
         Begin VB.Label Label31 
            Caption         =   "Ayuda de Doc."
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
            Left            =   1380
            TabIndex        =   119
            Top             =   390
            Width           =   1185
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   1
            Left            =   150
            Picture         =   "PrcCorrelativos.frx":0496
            Top             =   330
            Width           =   480
         End
         Begin VB.Label Label30 
            Caption         =   "[ F1 ]"
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
            Left            =   900
            TabIndex        =   118
            Top             =   390
            Width           =   435
         End
      End
      Begin VB.Frame Frame2 
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
         Height          =   885
         Index           =   0
         Left            =   390
         TabIndex        =   114
         Top             =   7815
         Width           =   2835
         Begin VB.Label Label31 
            Caption         =   "Ayuda de Doc."
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
            Index           =   2
            Left            =   1380
            TabIndex        =   116
            Top             =   390
            Width           =   1185
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   0
            Left            =   150
            Picture         =   "PrcCorrelativos.frx":08D8
            Top             =   330
            Width           =   480
         End
         Begin VB.Label Label30 
            Caption         =   "[ F1 ]"
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
            Left            =   900
            TabIndex        =   115
            Top             =   390
            Width           =   435
         End
      End
      Begin VB.Frame Frame4 
         Height          =   930
         Index           =   2
         Left            =   -71190
         TabIndex        =   100
         Top             =   6300
         Width           =   1980
         Begin VB.CommandButton cmdBotones 
            Caption         =   "&Acepta"
            Height          =   690
            Index           =   2
            Left            =   90
            Picture         =   "PrcCorrelativos.frx":0D1A
            Style           =   1  'Graphical
            TabIndex        =   84
            Top             =   180
            Width           =   870
         End
         Begin VB.CommandButton cmdBotones 
            Caption         =   "&Cancelar"
            Height          =   690
            Index           =   3
            Left            =   1080
            Picture         =   "PrcCorrelativos.frx":115C
            Style           =   1  'Graphical
            TabIndex        =   85
            Top             =   180
            Width           =   855
         End
      End
      Begin VB.Frame Fr1 
         BorderStyle     =   0  'None
         Height          =   4995
         Index           =   2
         Left            =   -74400
         TabIndex        =   74
         Top             =   1320
         Width           =   7635
         Begin VB.Frame Fr11 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Height          =   3075
            Left            =   150
            TabIndex        =   75
            Top             =   1140
            Width           =   7290
            Begin VB.ComboBox Combo6 
               Height          =   315
               Left            =   4470
               Style           =   2  'Dropdown List
               TabIndex        =   83
               Top             =   2130
               Width           =   1515
            End
            Begin MSMask.MaskEdBox MBox 
               Height          =   285
               Index           =   16
               Left            =   1590
               TabIndex        =   76
               Top             =   360
               Width           =   435
               _ExtentX        =   767
               _ExtentY        =   503
               _Version        =   393216
               Appearance      =   0
               MaxLength       =   2
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MBox 
               Height          =   285
               Index           =   17
               Left            =   2100
               TabIndex        =   77
               Top             =   360
               Width           =   585
               _ExtentX        =   1032
               _ExtentY        =   503
               _Version        =   393216
               Appearance      =   0
               MaxLength       =   3
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MBox 
               Height          =   285
               Index           =   18
               Left            =   2760
               TabIndex        =   78
               Top             =   360
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               _Version        =   393216
               Appearance      =   0
               MaxLength       =   8
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MBox 
               Height          =   285
               Index           =   19
               Left            =   5910
               TabIndex        =   79
               Top             =   330
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   503
               _Version        =   393216
               Appearance      =   0
               MaxLength       =   10
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MBox 
               Height          =   285
               Index           =   20
               Left            =   1590
               TabIndex        =   80
               Top             =   840
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   503
               _Version        =   393216
               Appearance      =   0
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MBox 
               Height          =   285
               Index           =   21
               Left            =   1560
               TabIndex        =   86
               Top             =   1260
               Width           =   1185
               _ExtentX        =   2090
               _ExtentY        =   503
               _Version        =   393216
               Appearance      =   0
               BackColor       =   -2147483648
               Enabled         =   0   'False
               MaxLength       =   11
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MBox 
               Height          =   285
               Index           =   22
               Left            =   1560
               TabIndex        =   81
               Top             =   1680
               Width           =   555
               _ExtentX        =   979
               _ExtentY        =   503
               _Version        =   393216
               Appearance      =   0
               MaxLength       =   3
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MBox 
               Height          =   285
               Index           =   23
               Left            =   1560
               TabIndex        =   82
               Top             =   2100
               Width           =   555
               _ExtentX        =   979
               _ExtentY        =   503
               _Version        =   393216
               Appearance      =   0
               MaxLength       =   3
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MBox 
               Height          =   285
               Index           =   28
               Left            =   1560
               TabIndex        =   112
               Top             =   2550
               Width           =   555
               _ExtentX        =   979
               _ExtentY        =   503
               _Version        =   393216
               Appearance      =   0
               MaxLength       =   2
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MBox 
               Height          =   285
               Index           =   29
               Left            =   4470
               TabIndex        =   113
               Top             =   2580
               Width           =   555
               _ExtentX        =   979
               _ExtentY        =   503
               _Version        =   393216
               Appearance      =   0
               MaxLength       =   2
               PromptChar      =   "_"
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Modo Venta"
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
               Index           =   29
               Left            =   2940
               TabIndex        =   109
               Top             =   2580
               Width           =   1305
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Forma de Pago"
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
               Index           =   28
               Left            =   240
               TabIndex        =   108
               Top             =   2580
               Width           =   1305
            End
            Begin VB.Label Label19 
               Caption         =   "Label7"
               Height          =   195
               Left            =   6750
               TabIndex        =   96
               Top             =   3090
               Width           =   525
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Documento"
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
               Index           =   22
               Left            =   270
               TabIndex        =   95
               Top             =   390
               Width           =   1095
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Cliente"
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
               Index           =   21
               Left            =   270
               TabIndex        =   94
               Top             =   900
               Width           =   615
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Fecha"
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
               Index           =   20
               Left            =   5190
               TabIndex        =   93
               Top             =   390
               Width           =   615
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Vendedor"
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
               Index           =   19
               Left            =   270
               TabIndex        =   92
               Top             =   1740
               Width           =   885
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Punto  Venta"
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
               Index           =   18
               Left            =   270
               TabIndex        =   91
               Top             =   2160
               Width           =   1155
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Estado  Dcmto"
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
               Index           =   17
               Left            =   2940
               TabIndex        =   90
               Top             =   2190
               Width           =   1305
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Ruc"
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
               Index           =   16
               Left            =   270
               TabIndex        =   89
               Top             =   1350
               Width           =   885
            End
            Begin VB.Label Label18 
               Appearance      =   0  'Flat
               BackColor       =   &H80000004&
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
               ForeColor       =   &H80000008&
               Height          =   285
               Left            =   2820
               TabIndex        =   88
               Top             =   840
               Width           =   4245
            End
            Begin VB.Label Label17 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000A&
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
               ForeColor       =   &H80000008&
               Height          =   285
               Left            =   2190
               TabIndex        =   87
               Top             =   1680
               Width           =   4875
            End
            Begin VB.Shape Shape4 
               BackColor       =   &H00800000&
               BackStyle       =   1  'Opaque
               Height          =   2865
               Left            =   120
               Shape           =   4  'Rounded Rectangle
               Top             =   120
               Width           =   7065
            End
            Begin VB.Shape Shape5 
               BackColor       =   &H00C0C0C0&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H0080FF80&
               FillColor       =   &H0000FF00&
               FillStyle       =   0  'Solid
               Height          =   3045
               Left            =   30
               Shape           =   4  'Rounded Rectangle
               Top             =   30
               Width           =   7245
            End
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000005&
            Index           =   11
            X1              =   60
            X2              =   7620
            Y1              =   4320
            Y2              =   4320
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00808080&
            Index           =   10
            X1              =   30
            X2              =   7620
            Y1              =   4290
            Y2              =   4290
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000003&
            Index           =   9
            X1              =   30
            X2              =   7620
            Y1              =   1080
            Y2              =   1080
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000005&
            Index           =   8
            X1              =   30
            X2              =   7590
            Y1              =   1110
            Y2              =   1110
         End
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   6150
            TabIndex        =   99
            Top             =   4530
            Width           =   1245
         End
         Begin VB.Label Label1 
            Caption         =   "IMPORTE"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   23
            Left            =   4650
            TabIndex        =   98
            Top             =   4530
            Width           =   1125
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            BackStyle       =   0  'Transparent
            Caption         =   "CONDICION DEL DOCUMENTO"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0FFC0&
            Height          =   435
            Index           =   2
            Left            =   240
            TabIndex        =   97
            Top             =   330
            Width           =   7065
         End
         Begin VB.Shape Shape6 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  'Opaque
            BorderStyle     =   6  'Inside Solid
            FillColor       =   &H00800000&
            FillStyle       =   0  'Solid
            Height          =   615
            Index           =   0
            Left            =   120
            Shape           =   4  'Rounded Rectangle
            Top             =   210
            Width           =   7275
         End
         Begin VB.Shape Shape6 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  'Opaque
            FillColor       =   &H0080FF80&
            FillStyle       =   0  'Solid
            Height          =   795
            Index           =   1
            Left            =   30
            Shape           =   4  'Rounded Rectangle
            Top             =   120
            Width           =   7485
         End
      End
      Begin VB.Frame Frame3 
         Height          =   645
         Left            =   -73260
         TabIndex        =   71
         Top             =   660
         Width           =   5505
         Begin VB.ComboBox Combo5 
            Height          =   315
            Left            =   900
            Style           =   2  'Dropdown List
            TabIndex        =   60
            Top             =   180
            Width           =   1425
         End
         Begin VB.CommandButton cBusca3 
            Caption         =   "&Buscar"
            Height          =   375
            Left            =   4260
            TabIndex        =   63
            Top             =   180
            Width           =   1095
         End
         Begin VB.TextBox aBusca 
            Height          =   285
            Index           =   4
            Left            =   2400
            MaxLength       =   3
            TabIndex        =   61
            Top             =   210
            Width           =   435
         End
         Begin VB.TextBox aBusca 
            Height          =   285
            Index           =   5
            Left            =   3060
            MaxLength       =   8
            TabIndex        =   62
            Top             =   210
            Width           =   885
         End
         Begin VB.Label Label16 
            Caption         =   "Tipo Doc."
            ForeColor       =   &H00800000&
            Height          =   225
            Left            =   180
            TabIndex        =   73
            Top             =   270
            Width           =   915
         End
         Begin VB.Label Label15 
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2940
            TabIndex        =   72
            Top             =   240
            Width           =   195
         End
      End
      Begin VB.Frame Frame4 
         Height          =   930
         Index           =   1
         Left            =   -71130
         TabIndex        =   70
         Top             =   6330
         Width           =   1980
         Begin VB.CommandButton cmdBotones 
            Caption         =   "&Cancelar"
            Height          =   690
            Index           =   1
            Left            =   1050
            Picture         =   "PrcCorrelativos.frx":159E
            Style           =   1  'Graphical
            TabIndex        =   52
            Top             =   180
            Width           =   855
         End
         Begin VB.CommandButton cmdBotones 
            Caption         =   "&Acepta"
            Height          =   690
            Index           =   0
            Left            =   90
            Picture         =   "PrcCorrelativos.frx":19E0
            Style           =   1  'Graphical
            TabIndex        =   51
            Top             =   180
            Width           =   870
         End
      End
      Begin VB.Frame Fr1 
         BorderStyle     =   0  'None
         Height          =   5085
         Index           =   1
         Left            =   -74400
         TabIndex        =   41
         Top             =   1320
         Width           =   7635
         Begin VB.Frame Fr10 
            BorderStyle     =   0  'None
            Height          =   3255
            Left            =   150
            TabIndex        =   42
            Top             =   1140
            Width           =   7380
            Begin VB.ComboBox Combo4 
               Height          =   315
               Left            =   4470
               Style           =   2  'Dropdown List
               TabIndex        =   50
               Top             =   2040
               Width           =   1515
            End
            Begin MSMask.MaskEdBox MBox 
               Height          =   285
               Index           =   8
               Left            =   1590
               TabIndex        =   43
               Top             =   360
               Width           =   435
               _ExtentX        =   767
               _ExtentY        =   503
               _Version        =   393216
               Appearance      =   0
               MaxLength       =   2
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MBox 
               Height          =   285
               Index           =   9
               Left            =   2100
               TabIndex        =   44
               Top             =   360
               Width           =   585
               _ExtentX        =   1032
               _ExtentY        =   503
               _Version        =   393216
               Appearance      =   0
               MaxLength       =   3
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MBox 
               Height          =   285
               Index           =   10
               Left            =   2760
               TabIndex        =   45
               Top             =   360
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               _Version        =   393216
               Appearance      =   0
               MaxLength       =   8
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MBox 
               Height          =   285
               Index           =   11
               Left            =   6030
               TabIndex        =   46
               Top             =   330
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   503
               _Version        =   393216
               Appearance      =   0
               MaxLength       =   10
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MBox 
               Height          =   285
               Index           =   12
               Left            =   1590
               TabIndex        =   47
               Top             =   840
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   503
               _Version        =   393216
               Appearance      =   0
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MBox 
               Height          =   285
               Index           =   13
               Left            =   1560
               TabIndex        =   53
               Top             =   1230
               Width           =   1185
               _ExtentX        =   2090
               _ExtentY        =   503
               _Version        =   393216
               Appearance      =   0
               BackColor       =   -2147483648
               Enabled         =   0   'False
               MaxLength       =   11
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MBox 
               Height          =   285
               Index           =   14
               Left            =   1560
               TabIndex        =   48
               Top             =   1650
               Width           =   555
               _ExtentX        =   979
               _ExtentY        =   503
               _Version        =   393216
               Appearance      =   0
               MaxLength       =   3
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MBox 
               Height          =   285
               Index           =   15
               Left            =   1560
               TabIndex        =   49
               Top             =   2070
               Width           =   555
               _ExtentX        =   979
               _ExtentY        =   503
               _Version        =   393216
               Appearance      =   0
               MaxLength       =   3
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MBox 
               Height          =   285
               Index           =   26
               Left            =   1590
               TabIndex        =   110
               Top             =   2490
               Width           =   555
               _ExtentX        =   979
               _ExtentY        =   503
               _Version        =   393216
               Appearance      =   0
               MaxLength       =   2
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MBox 
               Height          =   285
               Index           =   27
               Left            =   4470
               TabIndex        =   111
               Top             =   2520
               Width           =   555
               _ExtentX        =   979
               _ExtentY        =   503
               _Version        =   393216
               Appearance      =   0
               MaxLength       =   2
               PromptChar      =   "_"
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Modo Venta"
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
               Index           =   27
               Left            =   2910
               TabIndex        =   107
               Top             =   2550
               Width           =   1305
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Forma de Pago"
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
               Index           =   26
               Left            =   270
               TabIndex        =   106
               Top             =   2520
               Width           =   1305
            End
            Begin VB.Label Label13 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000A&
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
               ForeColor       =   &H80000008&
               Height          =   285
               Left            =   2190
               TabIndex        =   67
               Top             =   1650
               Width           =   4935
            End
            Begin VB.Label Label12 
               Appearance      =   0  'Flat
               BackColor       =   &H80000004&
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
               ForeColor       =   &H80000008&
               Height          =   285
               Left            =   2880
               TabIndex        =   66
               Top             =   840
               Width           =   4245
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Ruc"
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
               Index           =   14
               Left            =   270
               TabIndex        =   65
               Top             =   1290
               Width           =   885
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Estado  Dcmto"
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
               Index           =   13
               Left            =   2940
               TabIndex        =   64
               Top             =   2130
               Width           =   1305
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Punto  Venta"
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
               Index           =   12
               Left            =   270
               TabIndex        =   59
               Top             =   2100
               Width           =   1155
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Vendedor"
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
               Index           =   11
               Left            =   270
               TabIndex        =   58
               Top             =   1680
               Width           =   885
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Fecha"
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
               Index           =   10
               Left            =   5310
               TabIndex        =   57
               Top             =   390
               Width           =   615
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Cliente"
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
               Index           =   9
               Left            =   270
               TabIndex        =   56
               Top             =   900
               Width           =   615
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Documento"
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
               Index           =   8
               Left            =   300
               TabIndex        =   55
               Top             =   390
               Width           =   1095
            End
            Begin VB.Label Label11 
               Caption         =   "Label7"
               Height          =   195
               Left            =   6510
               TabIndex        =   54
               Top             =   3270
               Width           =   525
            End
            Begin VB.Shape Shape2 
               BackColor       =   &H00800000&
               BackStyle       =   1  'Opaque
               BorderStyle     =   6  'Inside Solid
               Height          =   2745
               Left            =   120
               Shape           =   4  'Rounded Rectangle
               Top             =   210
               Width           =   7125
            End
            Begin VB.Shape Shape3 
               BackColor       =   &H00C0C0C0&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H0080FF80&
               BorderStyle     =   6  'Inside Solid
               FillColor       =   &H0080FF80&
               FillStyle       =   0  'Solid
               Height          =   2925
               Left            =   30
               Shape           =   4  'Rounded Rectangle
               Top             =   120
               Width           =   7305
            End
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            BackStyle       =   0  'Transparent
            Caption         =   "CONDICION DEL DOCUMENTO"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0FFC0&
            Height          =   405
            Index           =   3
            Left            =   300
            TabIndex        =   101
            Top             =   360
            Width           =   7065
         End
         Begin VB.Label Label1 
            Caption         =   "IMPORTE"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   15
            Left            =   4650
            TabIndex        =   69
            Top             =   4620
            Width           =   1125
         End
         Begin VB.Label Label14 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   5880
            TabIndex        =   68
            Top             =   4620
            Width           =   1515
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000005&
            Index           =   7
            X1              =   60
            X2              =   7620
            Y1              =   1110
            Y2              =   1110
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000003&
            Index           =   6
            X1              =   60
            X2              =   7650
            Y1              =   1080
            Y2              =   1080
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00808080&
            Index           =   5
            X1              =   30
            X2              =   7620
            Y1              =   4410
            Y2              =   4410
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000005&
            Index           =   4
            X1              =   60
            X2              =   7620
            Y1              =   4440
            Y2              =   4440
         End
         Begin VB.Shape Shape6 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  'Opaque
            FillColor       =   &H00800000&
            FillStyle       =   0  'Solid
            Height          =   675
            Index           =   2
            Left            =   180
            Shape           =   4  'Rounded Rectangle
            Top             =   210
            Width           =   7305
         End
         Begin VB.Shape Shape6 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  'Opaque
            FillColor       =   &H0080FF80&
            FillStyle       =   0  'Solid
            Height          =   855
            Index           =   3
            Left            =   90
            Shape           =   4  'Rounded Rectangle
            Top             =   120
            Width           =   7485
         End
      End
      Begin VB.Frame Frame1 
         Height          =   645
         Left            =   -73260
         TabIndex        =   38
         Top             =   660
         Width           =   5505
         Begin VB.TextBox aBusca 
            Height          =   285
            Index           =   3
            Left            =   3060
            MaxLength       =   8
            TabIndex        =   30
            Top             =   210
            Width           =   885
         End
         Begin VB.TextBox aBusca 
            Height          =   285
            Index           =   2
            Left            =   2400
            MaxLength       =   3
            TabIndex        =   29
            Top             =   210
            Width           =   435
         End
         Begin VB.CommandButton cBusca2 
            Caption         =   "&Buscar"
            Height          =   375
            Left            =   4260
            TabIndex        =   31
            Top             =   180
            Width           =   1095
         End
         Begin VB.ComboBox Combo3 
            Height          =   315
            Left            =   900
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   180
            Width           =   1425
         End
         Begin VB.Label Label21 
            Caption         =   "Tipo Doc."
            ForeColor       =   &H00800000&
            Height          =   225
            Left            =   90
            TabIndex        =   123
            Top             =   240
            Width           =   915
         End
         Begin VB.Label Label10 
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2940
            TabIndex        =   40
            Top             =   240
            Width           =   195
         End
         Begin VB.Label Label9 
            Caption         =   "Tipo Doc."
            ForeColor       =   &H00800000&
            Height          =   105
            Left            =   30
            TabIndex        =   39
            Top             =   690
            Width           =   915
         End
      End
      Begin VB.Frame Fr2 
         Height          =   645
         Left            =   1740
         TabIndex        =   9
         Top             =   660
         Width           =   5505
         Begin VB.ComboBox Combo2 
            Height          =   315
            Left            =   900
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   180
            Width           =   1425
         End
         Begin VB.CommandButton cBusca 
            Caption         =   "&Buscar"
            Height          =   375
            Left            =   4260
            TabIndex        =   3
            Top             =   180
            Width           =   1095
         End
         Begin VB.TextBox aBusca 
            Height          =   285
            Index           =   0
            Left            =   2400
            MaxLength       =   3
            TabIndex        =   1
            Top             =   210
            Width           =   435
         End
         Begin VB.TextBox aBusca 
            Height          =   285
            Index           =   1
            Left            =   3090
            MaxLength       =   8
            TabIndex        =   2
            Top             =   210
            Width           =   885
         End
         Begin VB.Label Label5 
            Caption         =   "Tipo Doc."
            ForeColor       =   &H00800000&
            Height          =   225
            Left            =   150
            TabIndex        =   11
            Top             =   240
            Width           =   915
         End
         Begin VB.Label Label6 
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2940
            TabIndex        =   10
            Top             =   240
            Width           =   195
         End
      End
      Begin VB.Frame Fr1 
         BorderStyle     =   0  'None
         Height          =   6510
         Index           =   0
         Left            =   360
         TabIndex        =   7
         Top             =   1230
         Width           =   8115
         Begin VB.Frame Fr9 
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   0  'None
            Height          =   5655
            Left            =   30
            TabIndex        =   12
            Top             =   780
            Width           =   7920
            Begin VB.Frame FrameCaja 
               BackColor       =   &H0080FF80&
               Height          =   855
               Left            =   240
               TabIndex        =   135
               Top             =   3360
               Width           =   7575
               Begin TextFer.TxFer TxFernumero 
                  Height          =   315
                  Left            =   4080
                  TabIndex        =   136
                  Top             =   390
                  Width           =   1335
                  _ExtentX        =   2355
                  _ExtentY        =   556
                  Appearance      =   0
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
                  MaxLength       =   20
                  Text            =   ""
                  SaltarAlEnter   =   -1  'True
                  Valor           =   ""
                  TipoDato        =   1
                  SignoNegativo   =   0   'False
                  MarcarTextoAlEnfoque=   -1  'True
               End
               Begin TextFer.TxFer TxFerimporte 
                  Height          =   315
                  Left            =   6060
                  TabIndex        =   137
                  Top             =   390
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   556
                  Appearance      =   0
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
                  Text            =   ""
                  SaltarAlEnter   =   -1  'True
                  Valor           =   ""
                  TipoDato        =   1
                  NumeroDecimales =   2
                  Formato         =   "###,###.##"
                  MarcarTextoAlEnfoque=   -1  'True
               End
               Begin TextFer.TxFer TxFermoneda 
                  Height          =   315
                  Left            =   5550
                  TabIndex        =   138
                  Top             =   390
                  Width           =   495
                  _ExtentX        =   873
                  _ExtentY        =   556
                  Appearance      =   0
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
                  Text            =   ""
                  SaltarAlEnter   =   -1  'True
                  Valor           =   ""
                  TipoDato        =   1
                  NoCaracteres    =   "a-z,A-Z"
                  MarcarTextoAlEnfoque=   -1  'True
                  NoRangoCadena   =   -1  'True
               End
               Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuoperacion 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   139
                  Top             =   390
                  Width           =   1740
                  _ExtentX        =   3069
                  _ExtentY        =   556
                  XcodMaxLongitud =   11
                  xcodwith        =   200
                  NomTabla        =   "vt_conceptosdepago"
                  TituloAyuda     =   "Busqueda de Concepto de Pagos"
                  ListaCampos     =   "pagocodigo(1),pagodescripcion(1),pagoefectivo(1)"
                  XcodCampo       =   "pagocodigo"
                  XListCampo      =   "pagodescripcion"
                  ListaCamposDescrip=   "Codigo,Descripcion,efectivo"
                  ListaCamposText =   "pagocodigo,pagodescripcion,pagoefectivo"
               End
               Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayutipo 
                  Height          =   315
                  Left            =   1920
                  TabIndex        =   140
                  Top             =   390
                  Width           =   2100
                  _ExtentX        =   3704
                  _ExtentY        =   556
                  XcodMaxLongitud =   11
                  xcodwith        =   200
                  NomTabla        =   "vt_conceptostipodepago"
                  TituloAyuda     =   "Busqueda de Concepto de Pagos"
                  ListaCampos     =   "pagotipocodigo(1),pagotipodescripcion(1)"
                  XcodCampo       =   "pagotipocodigo"
                  XListCampo      =   "pagotipodescripcion"
                  ListaCamposDescrip=   "Codigo,Descripcion"
                  ListaCamposText =   "pagotipocodigo,pagotipodescripcion"
               End
               Begin VB.Label Label3d 
                  AutoSize        =   -1  'True
                  BackColor       =   &H80000009&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Importe"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H8000000D&
                  Height          =   195
                  Left            =   6480
                  TabIndex        =   145
                  Top             =   120
                  Width           =   705
               End
               Begin VB.Label Label3b 
                  AutoSize        =   -1  'True
                  BackColor       =   &H80000009&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Numero"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H8000000D&
                  Height          =   195
                  Left            =   4260
                  TabIndex        =   144
                  Top             =   120
                  Width           =   660
               End
               Begin VB.Label Label3c 
                  AutoSize        =   -1  'True
                  BackColor       =   &H80000009&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Moneda"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H8000000D&
                  Height          =   195
                  Left            =   5370
                  TabIndex        =   143
                  Top             =   120
                  Width           =   675
               End
               Begin VB.Label Label3a 
                  AutoSize        =   -1  'True
                  BackColor       =   &H80000009&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Tipo de tarjeta"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H8000000D&
                  Height          =   195
                  Left            =   1950
                  TabIndex        =   142
                  Top             =   120
                  Width           =   1260
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackColor       =   &H80000009&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Operacion"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H8000000D&
                  Height          =   195
                  Index           =   33
                  Left            =   150
                  TabIndex        =   141
                  Top             =   120
                  Width           =   855
               End
            End
            Begin VB.TextBox TextContacto 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1575
               MaxLength       =   100
               TabIndex        =   130
               Top             =   4725
               Width           =   6030
            End
            Begin VB.TextBox TxtEntrega 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1575
               MaxLength       =   70
               TabIndex        =   129
               Top             =   1620
               Width           =   5550
            End
            Begin MSComCtl2.DTPicker DTFecEnt 
               Height          =   315
               Left            =   4500
               TabIndex        =   125
               Top             =   1200
               Width           =   1515
               _ExtentX        =   2672
               _ExtentY        =   556
               _Version        =   393216
               Format          =   61603841
               CurrentDate     =   39739
            End
            Begin VB.ComboBox Combo1 
               Height          =   315
               Left            =   6030
               Style           =   2  'Dropdown List
               TabIndex        =   20
               Top             =   4275
               Width           =   1515
            End
            Begin MSMask.MaskEdBox MBox 
               Height          =   285
               Index           =   0
               Left            =   1590
               TabIndex        =   13
               Top             =   360
               Width           =   435
               _ExtentX        =   767
               _ExtentY        =   503
               _Version        =   393216
               Appearance      =   0
               MaxLength       =   2
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MBox 
               Height          =   285
               Index           =   1
               Left            =   2100
               TabIndex        =   14
               Top             =   360
               Width           =   585
               _ExtentX        =   1032
               _ExtentY        =   503
               _Version        =   393216
               Appearance      =   0
               MaxLength       =   3
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MBox 
               Height          =   285
               Index           =   2
               Left            =   2760
               TabIndex        =   15
               Top             =   360
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               _Version        =   393216
               Appearance      =   0
               MaxLength       =   8
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MBox 
               Height          =   285
               Index           =   3
               Left            =   5940
               TabIndex        =   16
               Top             =   330
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   503
               _Version        =   393216
               Appearance      =   0
               MaxLength       =   10
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MBox 
               Height          =   285
               Index           =   4
               Left            =   1590
               TabIndex        =   17
               Top             =   840
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   503
               _Version        =   393216
               Appearance      =   0
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MBox 
               Height          =   285
               Index           =   5
               Left            =   1560
               TabIndex        =   21
               Top             =   1230
               Width           =   1185
               _ExtentX        =   2090
               _ExtentY        =   503
               _Version        =   393216
               Appearance      =   0
               BackColor       =   -2147483648
               Enabled         =   0   'False
               MaxLength       =   11
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MBox 
               Height          =   285
               Index           =   6
               Left            =   1560
               TabIndex        =   18
               Top             =   2055
               Width           =   555
               _ExtentX        =   979
               _ExtentY        =   503
               _Version        =   393216
               Appearance      =   0
               MaxLength       =   3
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MBox 
               Height          =   285
               Index           =   7
               Left            =   1560
               TabIndex        =   19
               Top             =   2505
               Width           =   555
               _ExtentX        =   979
               _ExtentY        =   503
               _Version        =   393216
               Appearance      =   0
               MaxLength       =   3
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MBox 
               Height          =   285
               Index           =   24
               Left            =   1800
               TabIndex        =   104
               Top             =   4245
               Width           =   555
               _ExtentX        =   979
               _ExtentY        =   503
               _Version        =   393216
               Appearance      =   0
               MaxLength       =   2
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MBox 
               Height          =   285
               Index           =   25
               Left            =   5790
               TabIndex        =   105
               Top             =   3015
               Visible         =   0   'False
               Width           =   555
               _ExtentX        =   979
               _ExtentY        =   503
               _Version        =   393216
               Appearance      =   0
               MaxLength       =   2
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox TxtHor 
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "hh:mm AMPM"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   10250
                  SubFormatType   =   4
               EndProperty
               Height          =   285
               Left            =   6360
               TabIndex        =   126
               Top             =   2535
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   503
               _Version        =   393216
               MaxLength       =   5
               Format          =   "HH:mm"
               Mask            =   "##:##"
               PromptChar      =   "_"
            End
            Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayumodovta 
               Height          =   315
               Left            =   1800
               TabIndex        =   132
               Top             =   3000
               Width           =   3795
               _ExtentX        =   6694
               _ExtentY        =   556
               XcodMaxLongitud =   3
               xcodwith        =   200
               NomTabla        =   "vt_modoventa"
               TituloAyuda     =   "Ayuda de Modo de ventas"
               ListaCampos     =   "modovtacodigo(1),modovtadescripcion(1)"
               XcodCampo       =   "modovtacodigo"
               XListCampo      =   "modovtadescripcion"
               ListaCamposDescrip=   "Codigo,Descripcion"
               ListaCamposText =   "modovtacodigo,modovtadescripcion"
            End
            Begin VB.Label Label1 
               Caption         =   "IMPORTE"
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
               Height          =   195
               Index           =   6
               Left            =   4680
               TabIndex        =   134
               Top             =   5280
               Width           =   1125
            End
            Begin VB.Label Label8 
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
               ForeColor       =   &H000000C0&
               Height          =   315
               Left            =   6150
               TabIndex        =   133
               Top             =   5280
               Width           =   1515
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Contacto :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   32
               Left            =   225
               TabIndex        =   131
               Top             =   4770
               Width           =   975
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Dir. Entrega :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   31
               Left            =   270
               TabIndex        =   128
               Top             =   1665
               Width           =   1245
            End
            Begin VB.Label Label22 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Hora :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Left            =   5760
               TabIndex        =   127
               Top             =   2565
               Width           =   540
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Fecha Entrega :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   30
               Left            =   3030
               TabIndex        =   124
               Top             =   1260
               Width           =   1380
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Modo Venta"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   25
               Left            =   300
               TabIndex        =   103
               Top             =   2955
               Width           =   1305
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Forma de Pago"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   24
               Left            =   240
               TabIndex        =   102
               Top             =   4185
               Width           =   1305
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Documento"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   0
               Left            =   270
               TabIndex        =   36
               Top             =   390
               Width           =   1095
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Cliente"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   1
               Left            =   270
               TabIndex        =   35
               Top             =   900
               Width           =   615
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "F. Documento"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   2
               Left            =   4380
               TabIndex        =   34
               Top             =   390
               Width           =   1575
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Vendedor"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   3
               Left            =   270
               TabIndex        =   33
               Top             =   2055
               Width           =   885
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Punto  Venta"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   4
               Left            =   270
               TabIndex        =   32
               Top             =   2535
               Width           =   1155
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Estado  Dcmto"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   225
               Index           =   5
               Left            =   4500
               TabIndex        =   27
               Top             =   4335
               Width           =   1305
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Ruc"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   7
               Left            =   270
               TabIndex        =   26
               Top             =   1290
               Width           =   885
            End
            Begin VB.Label Label3 
               Appearance      =   0  'Flat
               BackColor       =   &H80000004&
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
               ForeColor       =   &H80000008&
               Height          =   285
               Left            =   2880
               TabIndex        =   25
               Top             =   840
               Width           =   4245
            End
            Begin VB.Label Label4 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000A&
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
               ForeColor       =   &H80000008&
               Height          =   285
               Left            =   2190
               TabIndex        =   23
               Top             =   2055
               Width           =   4935
            End
            Begin VB.Label Label7 
               Caption         =   "Label7"
               Height          =   120
               Left            =   6660
               TabIndex        =   37
               Top             =   3135
               Width           =   525
            End
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000005&
            Index           =   3
            X1              =   30
            X2              =   7590
            Y1              =   5730
            Y2              =   5730
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00808080&
            Index           =   2
            X1              =   0
            X2              =   7590
            Y1              =   6495
            Y2              =   6495
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000003&
            Index           =   1
            X1              =   30
            X2              =   7620
            Y1              =   720
            Y2              =   720
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000005&
            Index           =   0
            X1              =   30
            X2              =   7590
            Y1              =   750
            Y2              =   750
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            BackStyle       =   0  'Transparent
            Caption         =   "CONDICION DEL DOCUMENTO"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   315
            Index           =   0
            Left            =   240
            TabIndex        =   8
            Top             =   225
            Width           =   6075
         End
      End
      Begin VB.Frame Frame4 
         Height          =   930
         Index           =   0
         Left            =   3960
         TabIndex        =   6
         Top             =   7785
         Width           =   2100
         Begin VB.CommandButton cmdBotones 
            Caption         =   "&Acepta"
            Height          =   690
            Index           =   11
            Left            =   90
            Picture         =   "PrcCorrelativos.frx":1E22
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   180
            Width           =   870
         End
         Begin VB.CommandButton cmdBotones 
            Caption         =   "&Cancelar"
            Height          =   690
            Index           =   12
            Left            =   1080
            Picture         =   "PrcCorrelativos.frx":2264
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   180
            Width           =   855
         End
      End
   End
   Begin MSComctlLib.StatusBar Panel 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   4
      Top             =   9255
      Width           =   9405
      _ExtentX        =   16589
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
Attribute VB_Name = "PrcCorrelativos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adll As New dllgeneral.dll_general

Private Sub aBusca_GotFocus(Index As Integer)
   If Index Like "[01]" Then
     Fr9.Visible = False
     Fr1(0).Enabled = False
   ElseIf Index Like "[23]" Then
     Fr10.Visible = False
     Fr1(1).Enabled = False
   Else
     Fr11.Visible = False
     Fr1(2).Enabled = False
     
   End If

End Sub

Private Sub aBusca_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  Dim nsql As String
  
  If KeyCode = 112 Then  ' Ayuda de Productos
      Select Case SSTab1.Tab
        Case 0
             If adll.ComboDato(Combo2.Text) = g_tipobol Then
                 nsql = "CASE pedidocondicionfactura WHEN '1' THEN '*' ELSE '' END,pedidofechafact as Fecha,pedidonrofact as Boleta,clientecodigo as Codigo,clienterazonsocial as Cliente,pedidototneto as Total"
              ElseIf adll.ComboDato(Combo2.Text) = g_tipofac Then
                 nsql = "CASE pedidocondicionfactura WHEN '1' THEN '*' ELSE '' END,pedidofechafact as Fecha,pedidonrofact as Factura,clientecodigo as Codigo,clienterazonsocial as Cliente,pedidototneto as Total"
              Else
                 nsql = "CASE pedidocondicionfactura WHEN '1' THEN '*' ELSE '' END,pedidofechafact as Fecha,pedidonrofact as Documento,clientecodigo as Codigo,clienterazonsocial as Cliente,pedidototneto as Total"
              End If
         Case 1
             If adll.ComboDato(Combo3.Text) = g_tipobol Then
                 nsql = "CASE pedidocondicionfactura WHEN '1' THEN '*' ELSE '' END,pedidofechafact as Fecha,pedidonrofact as Boleta,clientecodigo as Codigo,clienterazonsocial as Cliente,pedidototneto as Total"
              ElseIf adll.ComboDato(Combo3.Text) = g_tipofac Then
                 nsql = "CASE pedidocondicionfactura WHEN '1' THEN '*' ELSE '' END,pedidofechafact as Fecha,pedidonrofact as Factura,clientecodigo as Codigo,clienterazonsocial as Cliente,pedidototneto as Total"
              Else
                 nsql = "CASE pedidocondicionfactura WHEN '1' THEN '*' ELSE '' END,pedidofechafact as Fecha,pedidonrofact as Documento,clientecodigo as Codigo,clienterazonsocial as Cliente,pedidototneto as Total"
              End If
         Case 2
                If adll.ComboDato(Combo5.Text) = g_tipoped Then
                   nsql = "CASE pedidoestado WHEN '1' THEN '*' ELSE '' END,pedidofecha as Fecha,pedidonumero as Pedido,clientecodigo as Codigo,clienterazonsocial as Cliente,pedidototneto as Total"
                Else
                   Exit Sub
                End If
       End Select
       Dim sfiltra(1 To 2, 1 To 2) As String
       sfiltra(1, 1) = "Cliente": sfiltra(1, 2) = "clienterazonsocial"
       sfiltra(2, 1) = "Ruc": sfiltra(2, 2) = "clienteruc"
       FrmAyuda.TipoForma = 2
       FrmAyuda.Bdata = "0"
       FrmAyuda.BConexion = VGCNx
       FrmAyuda.BTabla = "vt_pedido"
       FrmAyuda.BCampos = nsql
       Select Case SSTab1.Tab
         Case 0
                FrmAyuda.BOrden = "pedidonrofact"
                FrmAyuda.BCondi = "Pedidotipofac='" & adll.ComboDato(Combo2.Text) & "'"
          Case 1
                FrmAyuda.BOrden = "pedidonrofact"
                FrmAyuda.BCondi = "pedidotipofac='" & adll.ComboDato(Combo3.Text) & "' and pedidocondicionfactura='1'"
          Case Else
                FrmAyuda.BOrden = "pedidonumero"
                FrmAyuda.BCondi = ""
       End Select
       FrmAyuda.BFiltro = sfiltra
       FrmAyuda.Show 1
       If Index Like "[01]" Then
           aBusca(0) = Left(nAyuda, aBusca(0).MaxLength)
           aBusca(1) = Right(nAyuda, aBusca(1).MaxLength)
       ElseIf Index Like "[23]" Then
           aBusca(2) = Left(nAyuda, aBusca(2).MaxLength)
           aBusca(3) = Right(nAyuda, aBusca(3).MaxLength)
       ElseIf Index Like "[45]" Then
           aBusca(4) = Left(nAyuda, aBusca(4).MaxLength)
           aBusca(5) = Right(nAyuda, aBusca(5).MaxLength)
           
       End If
       nAyuda = "": nDetalle = ""
  End If
End Sub

Private Sub aBusca_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = 13 Then
    If Index Like "[135]" Then
      aBusca(Index) = Right("00000000000" & Trim(aBusca(Index)), aBusca(Index).MaxLength)
      If Index = 1 Then
         cBusca.SetFocus
      ElseIf Index = 3 Then
         cBusca2.SetFocus
      ElseIf Index = 5 Then
         cBusca3.SetFocus
      End If
    ElseIf Index Like "[024]" Then
      aBusca(Index) = Right("00000" & Trim(aBusca(Index)), aBusca(Index).MaxLength)
      SendKeys "{tab}"
    End If
  End If
End Sub

Private Sub cBusca_Click()
    Dim rs As New ADODB.Recordset
    Dim rs2 As New ADODB.Recordset
    Dim nsql As String
    
    Fr1(0).Enabled = True
    If Len(Trim(aBusca(0).Text)) > 0 And Len(Trim(aBusca(1).Text)) > 0 Then
    SQL = "select * from vt_pedido where empresacodigo='" & VGParametros.empresacodigo & "'"
    SQL = SQL & " and pedidotipofac='" & adll.ComboDato(Combo2.Text) & "'and pedidonrofact='" & Format(aBusca(0), "000") & Format(aBusca(1), "00000000") & "'"
        Set rs = VGCNx.Execute(SQL)
        If rs.RecordCount > 0 Then
           Fr9.Visible = True
           Label7 = Escadena(rs!pedidonumero)
           Label2(0) = IIf(rs!pedidocondicionfactura = "1", "ANULADO", "ACTIVO")
           DTFecEnt.Value = Format(rs!pedidofecha, "dd/mm/yyyy")
           MBox(0) = adll.ComboDato(Combo2.Text)
           MBox(1) = Left(Escadena(IIf(adll.ComboDato(Combo2.Text) = g_tipobol, rs!pedidonrofact, rs!pedidonrofact)), MBox(1).MaxLength)
           MBox(2) = Right(Escadena(IIf(adll.ComboDato(Combo2.Text) = g_tipobol, rs!pedidonrofact, rs!pedidonrofact)), MBox(2).MaxLength)
           If IsNull(rs!pedidofechafact) Then
              MBox(3) = "00/00/0000"
           Else
              MBox(3) = Format(rs!pedidofechafact, "dd/mm/yyyy")
           End If
           TxtEntrega.Text = Escadena(rs!pedidoentrega)
           MBox(4) = Escadena(rs!clientecodigo)
           Label3 = Escadena(rs!clienterazonsocial)
           MBox(5) = Escadena(rs!clienteruc)
           MBox(6) = Escadena(rs!vendedorcodigo)
           Set rs2 = VGCNx.Execute("select * from vt_vendedor")
           If rs2.RecordCount > 0 Then
               Label4 = Escadena(rs2!vendedornombres)
           Else
             Label4 = ""
           End If
           rs2.Close
           MBox(7) = Escadena(rs!puntovtacodigo)
           MBox(24) = Escadena(rs!formapagocodigo)
           MBox(25) = Escadena(rs!modovtacodigo)
           Ctr_Ayumodovta.xclave = MBox(25): Ctr_Ayumodovta.Ejecutar
           If IsNull(rs!pedidocondicionfactura) Then
               Combo1.ListIndex = 0
           Else
              Combo1.ListIndex = VerificaCombo(Combo1, rs!pedidocondicionfactura)
           End If
           Label8 = DatoMoneda(rs!pedidomoneda) & numero(rs!pedidototneto)
           TxtHor.Text = rs!horaentrega
           TextContacto.Text = rs!pedidoobserva
           If VGParamSistem.tesoreriaenlinea = 1 Then
              SQL = " select * from vt_pagosencaja where empresacodigo='" & VGParametros.empresacodigo & "'"
              SQL = SQL & " and pedidonumero='" & Label7 & "'"
              Set rs2 = VGCNx.Execute(SQL)
              If rs2.RecordCount > 0 Then
                 FrameCaja.Visible = True
                 Ctr_Ayuoperacion.xclave = rs2!pagocodigo: Ctr_Ayuoperacion.Ejecutar
                 Ctr_Ayutipo.xclave = rs2!pagotipocodigo: Ctr_Ayutipo.Ejecutar
                 TxFermoneda.valor = rs2!monedacodigo
                 TxFerimporte.valor = rs2!pagoimporte
              End If
              MBox(1).SetFocus
           End If
           MBox(1).SetFocus
        Else
          MsgBox "No existe Documento....Verifique!!", vbInformation, MsgTitle
          Limpiartexto aBusca, 0, 1
          Combo2.SetFocus
        End If
        rs.Close
    Else
        MsgBox "No existe Documento....Verifique!!", vbInformation, MsgTitle
    End If
    Set rs = Nothing
    Set rs2 = Nothing
End Sub

Private Sub cBusca_GotFocus()
   Fr9.Visible = False
   Fr1(0).Enabled = False
End Sub

Private Sub cBusca2_Click()
    Dim rs As New ADODB.Recordset
    Dim rs2 As New ADODB.Recordset
    Dim nsql As String
    
    Fr1(1).Enabled = True
    If Len(Trim(aBusca(2).Text)) > 0 And Len(Trim(aBusca(3).Text)) > 0 Then
        If adll.ComboDato(Combo3.Text) = g_tipobol Then
           nsql = "select * from vt_pedido where pedidonrofact='" & Trim(aBusca(2) & aBusca(3)) & "' and pedidotipofac='" & g_tipobol & "'"
        ElseIf adll.ComboDato(Combo3.Text) = g_tipofac Then
           nsql = "select * from vt_pedido where pedidonrofact='" & Trim(aBusca(2) & aBusca(3)) & "' and pedidotipofac='" & g_tipofac & "'"
        Else
           nsql = "select * from vt_pedido where pedidonrofact='" & Trim(aBusca(2) & aBusca(3)) & "' and pedidotipofac='" & adll.ComboDato(Combo3.Text) & "'"
        'Else
        '   Exit Sub
        End If
        Set rs = VGCNx.Execute(nsql)
        If rs.RecordCount > 0 Then
           Fr10.Visible = True
           Label11 = Escadena(rs!pedidonumero)
           Label2(1) = IIf(rs!pedidocondicionfactura = "1", "ANULADO", "ACTIVO")
           MBox(8) = adll.ComboDato(Combo3.Text)
           MBox(9) = Left(Escadena(IIf(adll.ComboDato(Combo3.Text) = g_tipobol, rs!pedidonrofact, rs!pedidonrofact)), MBox(1).MaxLength)
           MBox(10) = Right(Escadena(IIf(adll.ComboDato(Combo3.Text) = g_tipobol, rs!pedidonrofact, rs!pedidonrofact)), MBox(2).MaxLength)
           If IsNull(rs!pedidofechafact) Then
              MBox(11) = "00/00/0000"
           Else
              MBox(11) = Format(rs!pedidofechafact, "dd/mm/yyyy")
           End If
           MBox(12) = Escadena(rs!clientecodigo)
           Label12 = Escadena(rs!clienterazonsocial)
           MBox(13) = Escadena(rs!clienteruc)
           MBox(14) = Escadena(rs!vendedorcodigo)
           Set rs2 = VGCNx.Execute("select * from vt_vendedor")
           If rs2.RecordCount > 0 Then
               Label13 = Escadena(rs2!vendedornombres)
           Else
             Label13 = ""
           End If
           rs2.Close
           MBox(15) = Escadena(rs!puntovtacodigo)
           MBox(26) = Escadena(rs!formapagocodigo)
           MBox(27) = Escadena(rs!modovtacodigo)
           
           If IsNull(rs!pedidocondicionfactura) Then
               Combo4.ListIndex = 0
           Else
              Combo4.ListIndex = VerificaCombo(Combo4, rs!pedidocondicionfactura)
           End If
           Label14 = DatoMoneda(rs!pedidomoneda) & numero(rs!pedidototneto)
           MBox(9).SetFocus
        Else
          MsgBox "No existe Documento....Verifique!!", vbInformation, MsgTitle
          Limpiartexto aBusca, 2, 3
          Combo2.SetFocus
        End If
        rs.Close
    Else
        MsgBox "No existe Documento....Verifique!!", vbInformation, MsgTitle
    End If
    Set rs = Nothing
    Set rs2 = Nothing

End Sub

Private Sub cBusca3_Click()
    Dim rs As New ADODB.Recordset
    Dim rs2 As New ADODB.Recordset
    Dim nsql As String
    
    Fr1(2).Enabled = True
    If Len(Trim(aBusca(4).Text)) > 0 And Len(Trim(aBusca(5).Text)) > 0 Then
        If adll.ComboDato(Combo5.Text) = g_tipoped Then
           nsql = "select * from vt_pedido where pedidonumero='" & Trim(aBusca(4) & aBusca(5)) & "'"
        Else
           Exit Sub
        End If
        Set rs = VGCNx.Execute(nsql)
        If rs.RecordCount > 0 Then
           Fr11.Visible = True
           Label9 = Escadena(rs!pedidonumero)
           Label2(2) = IIf(rs!pedidoestado = "1", "ANULADO", "ACTIVO")
           MBox(16) = IIf(adll.ComboDato(Combo5.Text) = g_tipoped, g_tipoped, "PE")
           MBox(17) = Left(Escadena(IIf(adll.ComboDato(Combo5.Text) = g_tipoped, rs!pedidonumero, "000")), MBox(1).MaxLength)
           MBox(18) = Right(Escadena(IIf(adll.ComboDato(Combo5.Text) = g_tipoped, rs!pedidonumero, "000000000")), MBox(2).MaxLength)
           If IsNull(rs!pedidofecha) Then
              MBox(19) = "00/00/0000"
           Else
              MBox(19) = Format(rs!pedidofecha, "dd/mm/yyyy")
           End If
           MBox(20) = Escadena(rs!clientecodigo)
           Label18 = Escadena(rs!clienterazonsocial)
           MBox(21) = Escadena(rs!clienteruc)
           MBox(22) = Escadena(rs!vendedorcodigo)
           Set rs2 = VGCNx.Execute("select * from vt_vendedor")
           If rs2.RecordCount > 0 Then
               Label17 = Escadena(rs2!vendedornombres)
           Else
             Label17 = ""
           End If
           rs2.Close
           MBox(23) = Escadena(rs!puntovtacodigo)
           MBox(28) = Escadena(rs!formapagocodigo)
           MBox(29) = Escadena(rs!modovtacodigo)

           If IsNull(rs!pedidocondicionfactura) Then
               Combo6.ListIndex = 0
           Else
              Combo6.ListIndex = VerificaCombo(Combo6, rs!pedidocondicionfactura)
           End If
           Label20 = DatoMoneda(rs!pedidomoneda) & numero(rs!pedidototneto)
           MBox(17).SetFocus
        Else
          MsgBox "No existe Documento....Verifique!!", vbInformation, MsgTitle
          Limpiartexto aBusca, 4, 5
          Combo5.SetFocus
        End If
        rs.Close
    Else
        MsgBox "No existe Documento....Verifique!!", vbInformation, MsgTitle
    End If
    Set rs = Nothing
    Set rs2 = Nothing

End Sub

Private Sub cmdBotones_Click(Index As Integer)
   Dim nsql As String
   Dim rs As New ADODB.Recordset
   Dim acmd As New ADODB.Command
   On Error GoTo nerror
   
   Select Case Index
    Case 0      'anulados
       If Len(Trim(MBox(9))) = 0 Then
          MsgBox "Falta Ingresar No.Serie....Verifique!!", vbInformation, MsgTitle
          If Fr1(1).Enabled = True Then
             Call adll.Enfoquetexto(MBox(9))
          End If
          Exit Sub
       End If
       If Len(Trim(MBox(10))) = 0 Then
          MsgBox "Falta Ingresar No. Documento....Verifique!!", vbInformation, MsgTitle
          If Fr1(1).Enabled = True Then
             Call adll.Enfoquetexto(MBox(10))
          End If
          Exit Sub
       End If
       
       If Len(Trim(MBox(11))) = 0 Then
          MsgBox "Falta Ingresar Fecha Documento....Verifique!!", vbInformation, MsgTitle
          If Fr1(1).Enabled = True Then
             Call adll.Enfoquetexto(MBox(11))
          End If
          Exit Sub
       End If
       If Len(Trim(MBox(12))) = 0 Then
          MsgBox "Falta Ingresar Codigo de Cliente....Verifique!!", vbInformation, MsgTitle
          If Fr1(1).Enabled = True Then
             Call adll.Enfoquetexto(MBox(12))
          End If
          Exit Sub
       End If
       If Len(Trim(MBox(14))) = 0 Then
          MsgBox "Falta Ingresar Codigo de Vendedor....Verifique!!", vbInformation, MsgTitle
          If Fr1(1).Enabled = True Then
             Call adll.Enfoquetexto(MBox(14))
          End If
          Exit Sub
       End If
       
       If Len(Trim(MBox(26))) = 0 Then
          MsgBox "Falta Ingresar Forma de Pago....Verifique!!", vbInformation, MsgTitle
          If Fr1(1).Enabled = True Then
             Call adll.Enfoquetexto(MBox(26))
          End If
          Exit Sub
       End If
       If Len(Trim(MBox(27))) = 0 Then
          MsgBox "Falta Ingresar Modo de Venta....Verifique!!", vbInformation, MsgTitle
          If Fr1(1).Enabled = True Then
             Call adll.Enfoquetexto(MBox(27))
          End If
          Exit Sub
       End If
       
       If MsgBox("Desea Grabar los Cambios?", vbYesNo, MsgTitle) = vbYes Then
           Set acmd.ActiveConnection = VGgeneral
           acmd.CommandType = adCmdStoredProc
           acmd.CommandText = "vt_modificafactura_pro"
           acmd.Prepared = True
           If adll.ComboDato(Combo3.Text) = g_tipobol Then
               With acmd
                .Parameters("@base") = VGCNx.DefaultDatabase
                .Parameters("@tipo") = "1"
                .Parameters("@tabla") = "vt_pedido"
                .Parameters("@pedido") = Escadena(Label11)
                .Parameters("@tipodocu") = g_tipobol
                .Parameters("@numero") = Trim(MBox(9) & MBox(10))
                .Parameters("@cliente") = MBox(12)
                .Parameters("@ruc") = MBox(13)
                .Parameters("@vendedor") = MBox(14)
                .Parameters("@puntovta") = MBox(15)
                .Parameters("@fecha") = MBox(11)
                .Parameters("@condicion") = adll.ComboDato(Combo3.Text)
                .Parameters("@forma") = MBox(26)
                .Parameters("@modo") = MBox(27)
                .Parameters("@contacto") = TextContacto.Text
                
               End With
               acmd.Execute
               Set acmd = Nothing
               
               Set rs = VGCNx.Execute("select * from vt_pedido where pedidonumero='" & Escadena(Label11) & "'")
               If rs.RecordCount > 0 Then
                    VGCNx.Execute "Delete From vt_cargo Where documentocargo='" & g_tipobol & "' and cargonumdoc='" & Trim(MBox(9) & MBox(10)) & "'"
                    
                    Set acmd.ActiveConnection = VGgeneral
                    acmd.CommandType = adCmdStoredProc
                    acmd.CommandTimeout = 0
                    acmd.CommandText = "vt_ingresacargofactura_pro"
                    acmd.Prepared = True
                    With acmd
                        .Parameters("@base") = VGCNx.DefaultDatabase
                        .Parameters("@tipo") = "1"
                        .Parameters("@tabla") = "vt_cargo"
                        .Parameters("@tipodocu") = g_tipobol
                        .Parameters("@numero") = Trim(MBox(9) & MBox(10))
                        .Parameters("@cliente") = Escadena(MBox(12))
                        .Parameters("@vendedor") = Escadena(MBox(14))
                        .Parameters("@zona") = "00"
                        .Parameters("@apefecemi") = CDate(rs!pedidofechafact)
                        .Parameters("@moneda") = rs!pedidomoneda
                        .Parameters("@apeimppag") = rs!pedidototneto
                        .Parameters("@usuario") = g_usuario
                        .Parameters("@tipocambio") = rs!pedidotipcambio
                        .Parameters("@fechaact") = Date
                        .Parameters("@flagcancel") = "0"
                        .Parameters("@fechavenci") = CDate(rs!pedidofechafact + CDbl(rs!pedidodiaspago)) ''''
                        .Parameters("@cargoabono") = "C"
                    End With
                    acmd.Execute
                    Set acmd = Nothing
                End If
                rs.Close
                Set rs = Nothing
               
           ElseIf adll.ComboDato(Combo3.Text) = g_tipofac Then
               With acmd
                .Parameters("@base") = VGCNx.DefaultDatabase
                .Parameters("@tipo") = "1"
                .Parameters("@tabla") = "vt_pedido"
                .Parameters("@pedido") = Escadena(Label11)
                .Parameters("@tipodocu") = g_tipofac
                .Parameters("@numero") = Trim(MBox(9) & MBox(10))
                .Parameters("@cliente") = MBox(12)
                .Parameters("@ruc") = MBox(13)
                .Parameters("@vendedor") = MBox(14)
                .Parameters("@puntovta") = MBox(15)
                .Parameters("@fecha") = MBox(11)
                .Parameters("@condicion") = adll.ComboDato(Combo3.Text)
                .Parameters("@forma") = MBox(26)
                .Parameters("@modo") = MBox(27)
               End With
               acmd.Execute
               Set acmd = Nothing
               
               Set rs = VGCNx.Execute("select * from vt_pedido where pedidonumero='" & Escadena(Label11) & "'")
               If rs.RecordCount > 0 Then
                    VGCNx.Execute "Delete From vt_cargo Where documentocargo='" & g_tipofac & "' and cargonumdoc='" & Trim(MBox(9) & MBox(10)) & "'"
                    
                    Set acmd.ActiveConnection = VGgeneral
                    acmd.CommandType = adCmdStoredProc
                    acmd.CommandTimeout = 0
                    acmd.CommandText = "vt_ingresacargofactura_pro"
                    acmd.Prepared = True
                    With acmd
                        .Parameters("@base") = VGCNx.DefaultDatabase
                        .Parameters("@tipo") = "1"
                        .Parameters("@tabla") = "vt_cargo"
                        .Parameters("@tipodocu") = g_tipofac
                        .Parameters("@numero") = Trim(MBox(9) & MBox(10))
                        .Parameters("@cliente") = Escadena(MBox(12))
                        .Parameters("@vendedor") = Escadena(MBox(14))
                        .Parameters("@zona") = "00"
                        .Parameters("@apefecemi") = CDate(rs!pedidofechafact)
                        .Parameters("@moneda") = rs!pedidomoneda
                        .Parameters("@apeimppag") = rs!pedidototneto
                        .Parameters("@usuario") = g_usuario
                        .Parameters("@tipocambio") = rs!pedidotipcambio
                        .Parameters("@fechaact") = Date
                        .Parameters("@flagcancel") = "0"
                        .Parameters("@fechavenci") = CDate(rs!pedidofechafact + CDbl(rs!pedidodiaspago)) ''''
                        .Parameters("@cargoabono") = "C"
                    End With
                    acmd.Execute
                    Set acmd = Nothing
                End If
                rs.Close
                Set rs = Nothing
               
           Else
               With acmd
                .Parameters("@base") = VGCNx.DefaultDatabase
                .Parameters("@tipo") = "1"
                .Parameters("@tabla") = "vt_pedido"
                .Parameters("@pedido") = Escadena(Label11)
                .Parameters("@tipodocu") = MBox(8)
                .Parameters("@numero") = Trim(MBox(9) & MBox(10))
                .Parameters("@cliente") = MBox(12)
                .Parameters("@ruc") = MBox(13)
                .Parameters("@vendedor") = MBox(14)
                .Parameters("@puntovta") = MBox(15)
                .Parameters("@fecha") = MBox(11)
                .Parameters("@condicion") = adll.ComboDato(Combo3.Text)
                .Parameters("@forma") = MBox(26)
                .Parameters("@modo") = MBox(27)
               End With
               acmd.Execute
               Set acmd = Nothing
               
               Set rs = VGCNx.Execute("select * from vt_pedido where pedidonumero='" & Escadena(Label11) & "'")
               If rs.RecordCount > 0 Then
                    VGCNx.Execute "Delete From vt_cargo Where documentocargo='" & MBox(8) & "' and cargonumdoc='" & Trim(MBox(9) & MBox(10)) & "'"
                    
                    Set acmd.ActiveConnection = VGgeneral
                    acmd.CommandType = adCmdStoredProc
                    acmd.CommandTimeout = 0
                    acmd.CommandText = "vt_ingresacargofactura_pro"
                    acmd.Prepared = True
                    With acmd
                        .Parameters("@base") = VGCNx.DefaultDatabase
                        .Parameters("@tipo") = "1"
                        .Parameters("@tabla") = "vt_cargo"
                        .Parameters("@tipodocu") = MBox(8)
                        .Parameters("@numero") = Trim(MBox(9) & MBox(10))
                        .Parameters("@cliente") = Escadena(MBox(12))
                        .Parameters("@vendedor") = Escadena(MBox(14))
                        .Parameters("@zona") = "00"
                        .Parameters("@apefecemi") = CDate(rs!pedidofechafact)
                        .Parameters("@moneda") = rs!pedidomoneda
                        .Parameters("@apeimppag") = rs!pedidototneto
                        .Parameters("@usuario") = g_usuario
                        .Parameters("@tipocambio") = rs!pedidotipcambio
                        .Parameters("@fechaact") = Date
                        .Parameters("@flagcancel") = "0"
                        .Parameters("@fechavenci") = CDate(rs!pedidofechafact + CDbl(rs!pedidodiaspago)) ''''
                        .Parameters("@cargoabono") = "C"
                    End With
                    acmd.Execute
                    Set acmd = Nothing
                End If
                rs.Close
                Set rs = Nothing
               
           End If
           Label2(1) = "CONDICION DEL DOCUMENTO"
           Limpiartexto MBox, 8, 10
           MBox(3) = "00/00/0000"
           Limpiartexto MBox, 12, 15
           Limpiartexto aBusca, 2, 3
           aBusca(0).SetFocus
       End If
       Label2(1) = "CONDICION DEL DOCUMENTO"
       Limpiartexto MBox, 8, 10
       MBox(11) = "00/00/0000"
       Limpiartexto MBox, 12, 15
       Limpiartexto aBusca, 2, 3
       aBusca(2).SetFocus
       Fr10.Visible = False
       Fr1(1).Enabled = False
    Case 2   'Pedidos
       If Len(Trim(MBox(17))) = 0 Then
          MsgBox "Falta Ingresar No.Serie....Verifique!!", vbInformation, MsgTitle
          If Fr1(2).Enabled = True Then
             Call adll.Enfoquetexto(MBox(17))
          End If
          Exit Sub
       End If
       If Len(Trim(MBox(18))) = 0 Then
          MsgBox "Falta Ingresar No. Documento....Verifique!!", vbInformation, MsgTitle
          If Fr1(2).Enabled = True Then
             Call adll.Enfoquetexto(MBox(18))
          End If
          Exit Sub
       End If
       
       If Len(Trim(MBox(19))) = 0 Then
          MsgBox "Falta Ingresar Fecha Documento....Verifique!!", vbInformation, MsgTitle
          If Fr1(2).Enabled = True Then
             Call adll.Enfoquetexto(MBox(19))
          End If
          Exit Sub
       End If
       If Len(Trim(MBox(20))) = 0 Then
          MsgBox "Falta Ingresar Codigo de Cliente....Verifique!!", vbInformation, MsgTitle
          If Fr1(2).Enabled = True Then
             Call adll.Enfoquetexto(MBox(20))
          End If
          Exit Sub
       End If
       If Len(Trim(MBox(22))) = 0 Then
          MsgBox "Falta Ingresar Codigo de Vendedor....Verifique!!", vbInformation, MsgTitle
          If Fr1(2).Enabled = True Then
             Call adll.Enfoquetexto(MBox(22))
          End If
          Exit Sub
       End If
       If Len(Trim(MBox(28))) = 0 Then
          MsgBox "Falta Ingresar Forma de Pago....Verifique!!", vbInformation, MsgTitle
          If Fr1(2).Enabled = True Then
             Call adll.Enfoquetexto(MBox(28))
          End If
          Exit Sub
       End If
       If Len(Trim(MBox(29))) = 0 Then
          MsgBox "Falta Ingresar Modo de Venta....Verifique!!", vbInformation, MsgTitle
          If Fr1(2).Enabled = True Then
             Call adll.Enfoquetexto(MBox(29))
          End If
          Exit Sub
       End If
       
       
       If MsgBox("Desea Grabar los Cambios?", vbYesNo, MsgTitle) = vbYes Then
           Set acmd.ActiveConnection = VGgeneral
           acmd.CommandType = adCmdStoredProc
           acmd.CommandText = "vt_modificafactura_pro"
           acmd.Prepared = True
           If adll.ComboDato(Combo5.Text) = g_tipoped Then
               With acmd
                .Parameters("@base") = VGCNx.DefaultDatabase
                .Parameters("@tipo") = "3"
                .Parameters("@tabla") = "vt_pedido"
                .Parameters("@pedido") = Escadena(Label9)
                .Parameters("@tipodocu") = g_tipoped
                .Parameters("@numero") = Trim(MBox(17) & MBox(18))
                .Parameters("@cliente") = MBox(20)
                .Parameters("@ruc") = MBox(21)
                .Parameters("@vendedor") = MBox(22)
                .Parameters("@puntovta") = MBox(15)
                .Parameters("@fecha") = MBox(19)
                .Parameters("@condicion") = adll.ComboDato(Combo6.Text)
                .Parameters("@forma") = MBox(28)
                .Parameters("@modo") = MBox(29)
               End With
               acmd.Execute
               Set acmd = Nothing
               
           End If
           Label2(2) = "CONDICION DEL DOCUMENTO"
           Limpiartexto MBox, 16, 18
           MBox(19) = "00/00/0000"
           Limpiartexto MBox, 20, 23
           Limpiartexto aBusca, 4, 5
           aBusca(4).SetFocus
       End If
       Label2(2) = "CONDICION DEL DOCUMENTO"
       Limpiartexto MBox, 16, 18
       MBox(19) = "00/00/0000"
       Limpiartexto MBox, 20, 23
       Limpiartexto aBusca, 4, 5
       aBusca(4).SetFocus
       Fr11.Visible = False
       Fr1(2).Enabled = False
    
    Case 11
       If Len(Trim(MBox(1))) = 0 Then
          MsgBox "Falta Ingresar No.Serie....Verifique!!", vbInformation, MsgTitle
          If Fr1(0).Enabled = True Then
             Call adll.Enfoquetexto(MBox(1))
          End If
          Exit Sub
       End If
       If Len(Trim(MBox(2))) = 0 Then
          MsgBox "Falta Ingresar No. Documento....Verifique!!", vbInformation, MsgTitle
          If Fr1(0).Enabled = True Then
             Call adll.Enfoquetexto(MBox(2))
          End If
          Exit Sub
       End If
       
       If Len(Trim(MBox(3))) = 0 Then
          MsgBox "Falta Ingresar Fecha Documento....Verifique!!", vbInformation, MsgTitle
          If Fr1(0).Enabled = True Then
             Call adll.Enfoquetexto(MBox(3))
          End If
          Exit Sub
       End If
       If Len(Trim(MBox(4))) = 0 Then
          MsgBox "Falta Ingresar Codigo de Cliente....Verifique!!", vbInformation, MsgTitle
          If Fr1(0).Enabled = True Then
             Call adll.Enfoquetexto(MBox(4))
          End If
          Exit Sub
       End If
       If Len(Trim(MBox(6))) = 0 Then
          MsgBox "Falta Ingresar Codigo de Vendedor....Verifique!!", vbInformation, MsgTitle
          If Fr1(0).Enabled = True Then
             Call adll.Enfoquetexto(MBox(5))
          End If
          Exit Sub
       End If
       If Len(Trim(MBox(24))) = 0 Then
          MsgBox "Falta Ingresar Forma de Pago....Verifique!!", vbInformation, MsgTitle
          If Fr1(0).Enabled = True Then
             Call adll.Enfoquetexto(MBox(24))
          End If
          Exit Sub
       End If
       If Len(Trim(MBox(25))) = 0 Then
          MsgBox "Falta Ingresar Modo de Venta....Verifique!!", vbInformation, MsgTitle
          If Fr1(0).Enabled = True Then
             Call adll.Enfoquetexto(MBox(25))
          End If
          Exit Sub
       End If
       
       If MsgBox("Desea Grabar los Cambios?", vbYesNo, MsgTitle) = vbYes Then
           Set acmd.ActiveConnection = VGgeneral
           acmd.CommandType = adCmdStoredProc
           acmd.CommandText = "vt_modificafactura_pro"
           acmd.Prepared = True
           With acmd
                .Parameters("@base") = VGCNx.DefaultDatabase
                .Parameters("@tipo") = "1"
                .Parameters("@tabla") = "vt_pedido"
                .Parameters("@pedido") = Escadena(Label7)
                .Parameters("@tipodocu") = MBox(0)
                .Parameters("@numero") = Trim(MBox(1) & MBox(2))
                .Parameters("@cliente") = MBox(4)
                .Parameters("@ruc") = MBox(5)
                .Parameters("@vendedor") = MBox(6)
                .Parameters("@puntovta") = MBox(7)
                .Parameters("@fecha") = MBox(3)
                .Parameters("@condicion") = adll.ComboDato(Combo1.Text)
                .Parameters("@forma") = MBox(24)
                .Parameters("@modo") = MBox(25)
                .Parameters("@entrega") = DTFecEnt.Value
                .Parameters("@hora") = TxtHor.Text
                .Parameters("@direntrega") = TxtEntrega.Text
                .Parameters("@empresa") = VGParametros.empresacodigo
                .Parameters("@contacto") = TextContacto.Text
                
               End With
               acmd.Execute
               Set acmd = Nothing
               
               SQL = "select * from vt_modoventa a inner join vt_pedido b on a.modovtacodigo=b.modovtacodigo "
               SQL = SQL & " where empresacodigo='" & VGParametros.empresacodigo & "' and pedidonumero='" & Escadena(Label7) & "'"
               SQL = SQL & " and a.modovtaactctacte=1 "
               Set rs = VGCNx.Execute(SQL)
               VGCNx.Execute "Delete From vt_cargo Where empresacodigo='" & VGParametros.empresacodigo & "' and  documentocargo='" & MBox(0) & "' and cargonumdoc='" & Trim(MBox(1) & MBox(2)) & "'"
               If rs.RecordCount > 0 Then
                    Set acmd.ActiveConnection = VGgeneral
                    acmd.CommandType = adCmdStoredProc
                    acmd.CommandTimeout = 0
                    acmd.CommandText = "vt_ingresacargofactura_pro"
                    acmd.Prepared = True
                    With acmd
                        .Parameters("@base") = VGCNx.DefaultDatabase
                        .Parameters("@tipo") = "1"
                        .Parameters("@tabla") = "vt_cargo"
                        .Parameters("@tipodocu") = MBox(0)
                        .Parameters("@numero") = Trim(MBox(1) & MBox(2))
                        .Parameters("@cliente") = Escadena(MBox(4))
                        .Parameters("@vendedor") = Escadena(MBox(6))
                        .Parameters("@zona") = "00"
                        .Parameters("@apefecemi") = CDate(rs!pedidofechafact)
                        .Parameters("@moneda") = rs!pedidomoneda
                        .Parameters("@apeimppag") = rs!pedidototneto
                        .Parameters("@usuario") = g_usuario
                        .Parameters("@tipocambio") = rs!pedidotipcambio
                        .Parameters("@fechaact") = Date
                        .Parameters("@flagcancel") = "0"
                        .Parameters("@fechavenci") = CDate(rs!pedidofechafact + CDbl(rs!pedidodiaspago)) ''''
                        .Parameters("@cargoabono") = "C"
                        .Parameters("@empresa") = VGParametros.empresacodigo
                    End With
                    acmd.Execute
                    Set acmd = Nothing
                End If
                rs.Close
                Set rs = Nothing
                              
           End If
           aBusca(0) = "'"
           aBusca(1) = ""
          Combo2.SetFocus
       Fr9.Visible = False
       Fr1(0).Enabled = False
    Case 12
        Fr9.Visible = False
        Fr1(0).Enabled = False
        Unload Me
    Case 1
        Fr10.Visible = False
        Fr1(1).Enabled = False
        Unload Me
        
    Case 3
        Fr11.Visible = False
        Fr1(2).Enabled = False
        Unload Me
   End Select
   Set acmd = Nothing
   
nerror:
   If Err Then
      MsgBox Err.Number & "-" & Err.Description
      MsgBox "Inconsistencia de Datos Referenciales ó Datos Incompletos", vbInformation, MsgTitle
      Err = 0
      Exit Sub
      Resume
   End If
End Sub





Private Sub Combo2_KeyPress(KeyAscii As Integer)
  Seguir Combo2, KeyAscii
End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
   Seguir Combo3, KeyAscii
End Sub

Private Sub Combo5_KeyPress(KeyAscii As Integer)
  Seguir Combo5, KeyAscii
End Sub

Private Sub Form_Load()
   MostrarForm Me, "C"
   Fr9.Visible = False
   Fr10.Visible = False
   Fr11.Visible = False
   
   Call Ctr_Ayumodovta.Conexion(VGCNx)
   Call Ctr_Ayuoperacion.Conexion(VGCNx)
   Call Ctr_Ayutipo.Conexion(VGCNx)
   
   If VgModificar = 1 Then
     Ctr_Ayumodovta.Visible = True
     MBox(24).Visible = True
     Combo1.Visible = True
   Else
     Ctr_Ayumodovta.Visible = True
     MBox(24).Visible = True
     Combo1.Visible = True
   End If
   Fr1(0).Enabled = False
   Fr1(1).Enabled = False
   Fr1(2).Enabled = False
   Call CargaOpcion
   SSTab1.Tab = 0
End Sub

Public Sub CargaOpcion()
   
   Call CargarTipo(Combo1, 1)
   
   Call CargarTipo(Combo2, 2)
   
   Call CargarTipo(Combo3, 2)

   Call CargarTipo(Combo4, 1)

   Combo5.Clear
   Combo5.AddItem g_tipoped & "-PEDIDO"
   Combo5.ListIndex = 0

   Call CargarTipo(Combo6, 1)

End Sub


Private Sub MBox_Change(Index As Integer)
  Select Case Index
    Case 4
      If Len(Trim(MBox(Index))) = 0 Then
         Label3 = ""
      End If
    Case 12
      If Len(Trim(MBox(Index))) = 0 Then
         Label12 = ""
      End If
    Case 5
      If Len(Trim(MBox(Index))) = 0 Then
         Label4 = ""
      End If
    Case 14
      If Len(Trim(MBox(Index))) = 0 Then
         Label13 = ""
      End If
   Case 20
      If Len(Trim(MBox(Index))) = 0 Then
         Label18 = ""
      End If
    Case 22
      If Len(Trim(MBox(Index))) = 0 Then
         Label17 = ""
      End If
      
    
 End Select
  
End Sub


Private Sub MBox_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then
     If Index = 4 Then
       MBox(6).SetFocus
     ElseIf Index = 12 Then
       MBox(14).SetFocus
     ElseIf Index = 20 Then
       MBox(22).SetFocus
    Else
       SendKeys "{tab}"
     End If
  ElseIf KeyCode = 112 Then
   If Index = 4 Or Index = 12 Or Index = 20 Then
       Dim sfiltra(1 To 3, 1 To 2) As String
       sfiltra(1, 1) = "Codigo": sfiltra(1, 2) = "clientecodigo"
       sfiltra(2, 1) = "Cliente": sfiltra(2, 2) = "clienterazonsocial"
       sfiltra(3, 1) = "Ruc": sfiltra(3, 2) = "clienteruc"
       FrmAyuda.TipoForma = 1
       FrmAyuda.BConexion = VGCNx
       FrmAyuda.BTabla = "vt_cliente"
       FrmAyuda.BCampos = "Clientecodigo as Codigo,Clienterazonsocial as Descripcion,Clienteruc as Ruc"
       FrmAyuda.BOrden = "clienterazonsocial"
       FrmAyuda.BCondi = ""
       FrmAyuda.BFiltro = sfiltra
       FrmAyuda.Show 1
       If Index = 4 Then
         MBox(4) = nAyuda
         Label3 = nDetalle
       ElseIf Index = 12 Then
         MBox(12) = nAyuda
         Label12 = nDetalle
       Else
         MBox(20) = nAyuda
         Label18 = nDetalle
       End If
       nAyuda = "": nDetalle = ""
   ElseIf Index = 6 Or Index = 14 Or Index = 22 Then
       Dim gfiltra(1 To 2, 1 To 2) As String
       gfiltra(1, 1) = "Codigo": gfiltra(1, 2) = "vendedorcodigo"
       gfiltra(2, 1) = "Vendedor": gfiltra(2, 2) = "vendedornombres"
       FrmAyuda.TipoForma = 1
       FrmAyuda.BConexion = VGCNx
       FrmAyuda.BTabla = "vt_vendedor"
       FrmAyuda.BCampos = "vendedorcodigo as Codigo,vendedornombres as Descripcion"
       FrmAyuda.BOrden = "vendedorcodigo"
       FrmAyuda.BCondi = ""
       FrmAyuda.BFiltro = gfiltra
       FrmAyuda.Show 1
       If Index = 6 Then
        MBox(6) = nAyuda
        Label4 = nDetalle
      ElseIf Index = 14 Then
        MBox(14) = nAyuda
        Label13 = nDetalle
      Else
        MBox(22) = nAyuda
        Label17 = nDetalle
      End If
      nAyuda = "": nDetalle = ""
 
   End If
  End If
End Sub

Private Sub MBox_LostFocus(Index As Integer)
  Dim rs As New ADODB.Recordset
  
  If Index = 4 Or Index = 12 Or Index = 20 Then
     If adll.VerificaDatoExistente(VGCNx, "select * from vt_cliente where clientecodigo='" & MBox(Index) & "'") = 0 And Len(Trim(MBox(Index))) > 0 Then
         MsgBox "No existe ese cliente...Verifique!!!", vbInformation, MsgTitle
         Call adll.Enfoquetexto(MBox(Index))
         Exit Sub
     Else
       Set rs = VGCNx.Execute("select * from vt_cliente where clientecodigo='" & MBox(Index) & "'")
       If rs.RecordCount > 0 Then
            If Index = 4 Then
              Label3 = Escadena(rs!clienterazonsocial)
            ElseIf Index = 12 Then
              Label12 = Escadena(rs!clienterazonsocial)
            Else
              Label18 = Escadena(rs!clienterazonsocial)
            End If
       End If
       rs.Close
     End If
  ElseIf Index = 6 Or Index = 14 Or Index = 22 Then
     If adll.VerificaDatoExistente(VGCNx, "select * from vt_vendedor where vendedorcodigo='" & MBox(Index) & "'") = 0 And Len(Trim(MBox(Index))) > 0 Then
         MsgBox "No existe ese Vendedor...Verifique!!!", vbInformation, MsgTitle
         Call adll.Enfoquetexto(MBox(Index))
         Exit Sub
     Else
        Set rs = VGCNx.Execute("select * from vt_vendedor where vendedorcodigo='" & MBox(Index) & "'")
        If rs.RecordCount > 0 Then
             If Index = 6 Then
                Label4 = Escadena(rs!vendedornombres)
              ElseIf Index = 14 Then
                Label13 = Escadena(rs!vendedornombres)
              Else
                Label17 = Escadena(rs!vendedornombres)
              End If
          End If
          rs.Close
     End If
  End If
  Set rs = Nothing
End Sub

