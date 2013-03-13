VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{4D137D9C-00A6-4458-9B46-7E95DB76D55B}#9.0#0"; "TextFer.ocx"
Begin VB.Form Frmcorrecciondatosgen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Procesos de Correlativos"
   ClientHeight    =   9840
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10575
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9840
   ScaleWidth      =   10575
   Begin TabDlg.SSTab SSTab1 
      Height          =   9240
      Left            =   240
      TabIndex        =   5
      Top             =   210
      Width           =   10125
      _ExtentX        =   17859
      _ExtentY        =   16298
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      ForeColor       =   8388608
      TabCaption(0)   =   "Documento Venta"
      TabPicture(0)   =   "FrmCorrecciondatosgen.frx":0000
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
      TabPicture(1)   =   "FrmCorrecciondatosgen.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(1)=   "Fr1(1)"
      Tab(1).Control(2)=   "Frame4(1)"
      Tab(1).Control(3)=   "Frame2(1)"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Ingresa No. Pedido"
      TabPicture(2)   =   "FrmCorrecciondatosgen.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame3"
      Tab(2).Control(1)=   "Fr1(2)"
      Tab(2).Control(2)=   "Frame4(2)"
      Tab(2).Control(3)=   "Frame2(2)"
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
         TabIndex        =   118
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
            TabIndex        =   120
            Top             =   390
            Width           =   1185
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   2
            Left            =   150
            Picture         =   "FrmCorrecciondatosgen.frx":0054
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
            TabIndex        =   119
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
         TabIndex        =   115
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
            TabIndex        =   117
            Top             =   390
            Width           =   1185
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   1
            Left            =   150
            Picture         =   "FrmCorrecciondatosgen.frx":0496
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
            TabIndex        =   116
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
         TabIndex        =   112
         Top             =   8295
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
            TabIndex        =   114
            Top             =   390
            Width           =   1185
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   0
            Left            =   150
            Picture         =   "FrmCorrecciondatosgen.frx":08D8
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
            TabIndex        =   113
            Top             =   390
            Width           =   435
         End
      End
      Begin VB.Frame Frame4 
         Height          =   930
         Index           =   2
         Left            =   -71190
         TabIndex        =   98
         Top             =   6300
         Width           =   1980
         Begin VB.CommandButton cmdBotones 
            Caption         =   "&Acepta"
            Height          =   690
            Index           =   2
            Left            =   90
            Picture         =   "FrmCorrecciondatosgen.frx":0D1A
            Style           =   1  'Graphical
            TabIndex        =   82
            Top             =   180
            Width           =   870
         End
         Begin VB.CommandButton cmdBotones 
            Caption         =   "&Cancelar"
            Height          =   690
            Index           =   3
            Left            =   1080
            Picture         =   "FrmCorrecciondatosgen.frx":115C
            Style           =   1  'Graphical
            TabIndex        =   83
            Top             =   180
            Width           =   855
         End
      End
      Begin VB.Frame Fr1 
         BorderStyle     =   0  'None
         Height          =   4995
         Index           =   2
         Left            =   -74400
         TabIndex        =   72
         Top             =   1320
         Width           =   7635
         Begin VB.Frame Fr11 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Height          =   3075
            Left            =   150
            TabIndex        =   73
            Top             =   1140
            Width           =   7290
            Begin VB.ComboBox Combo6 
               Height          =   315
               Left            =   4470
               Style           =   2  'Dropdown List
               TabIndex        =   81
               Top             =   2130
               Width           =   1515
            End
            Begin MSMask.MaskEdBox MBox 
               Height          =   285
               Index           =   16
               Left            =   1590
               TabIndex        =   74
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
               TabIndex        =   75
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
               TabIndex        =   76
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
               TabIndex        =   77
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
               TabIndex        =   78
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
               TabIndex        =   84
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
               TabIndex        =   79
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
               TabIndex        =   80
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
               TabIndex        =   110
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
               TabIndex        =   111
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
               TabIndex        =   107
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
               TabIndex        =   106
               Top             =   2580
               Width           =   1305
            End
            Begin VB.Label Label19 
               Caption         =   "Label7"
               Height          =   195
               Left            =   6750
               TabIndex        =   94
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
               TabIndex        =   93
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
               TabIndex        =   92
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
               TabIndex        =   91
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
               TabIndex        =   90
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
               TabIndex        =   89
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
               TabIndex        =   88
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
               TabIndex        =   87
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
               TabIndex        =   86
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
               TabIndex        =   85
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
            TabIndex        =   97
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
            TabIndex        =   96
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
            TabIndex        =   95
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
         TabIndex        =   69
         Top             =   660
         Width           =   5505
         Begin VB.ComboBox Combo5 
            Height          =   315
            Left            =   900
            Style           =   2  'Dropdown List
            TabIndex        =   58
            Top             =   180
            Width           =   1425
         End
         Begin VB.CommandButton cBusca3 
            Caption         =   "&Buscar"
            Height          =   375
            Left            =   4260
            TabIndex        =   61
            Top             =   180
            Width           =   1095
         End
         Begin VB.TextBox aBusca 
            Height          =   285
            Index           =   4
            Left            =   2400
            MaxLength       =   3
            TabIndex        =   59
            Top             =   210
            Width           =   435
         End
         Begin VB.TextBox aBusca 
            Height          =   285
            Index           =   5
            Left            =   3060
            MaxLength       =   8
            TabIndex        =   60
            Top             =   210
            Width           =   885
         End
         Begin VB.Label Label16 
            Caption         =   "Tipo Doc."
            ForeColor       =   &H00800000&
            Height          =   225
            Left            =   180
            TabIndex        =   71
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
            TabIndex        =   70
            Top             =   240
            Width           =   195
         End
      End
      Begin VB.Frame Frame4 
         Height          =   930
         Index           =   1
         Left            =   -71130
         TabIndex        =   68
         Top             =   6330
         Width           =   1980
         Begin VB.CommandButton cmdBotones 
            Caption         =   "&Cancelar"
            Height          =   690
            Index           =   1
            Left            =   1050
            Picture         =   "FrmCorrecciondatosgen.frx":159E
            Style           =   1  'Graphical
            TabIndex        =   50
            Top             =   180
            Width           =   855
         End
         Begin VB.CommandButton cmdBotones 
            Caption         =   "&Acepta"
            Height          =   690
            Index           =   0
            Left            =   90
            Picture         =   "FrmCorrecciondatosgen.frx":19E0
            Style           =   1  'Graphical
            TabIndex        =   49
            Top             =   180
            Width           =   870
         End
      End
      Begin VB.Frame Fr1 
         BorderStyle     =   0  'None
         Height          =   5085
         Index           =   1
         Left            =   -74400
         TabIndex        =   39
         Top             =   1320
         Width           =   7635
         Begin VB.Frame Fr10 
            BorderStyle     =   0  'None
            Height          =   3255
            Left            =   150
            TabIndex        =   40
            Top             =   1140
            Width           =   7380
            Begin VB.ComboBox Combo4 
               Height          =   315
               Left            =   4470
               Style           =   2  'Dropdown List
               TabIndex        =   48
               Top             =   2040
               Width           =   1515
            End
            Begin MSMask.MaskEdBox MBox 
               Height          =   285
               Index           =   8
               Left            =   1590
               TabIndex        =   41
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
               TabIndex        =   42
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
               TabIndex        =   43
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
               TabIndex        =   44
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
               TabIndex        =   45
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
               TabIndex        =   51
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
               TabIndex        =   46
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
               TabIndex        =   47
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
               TabIndex        =   108
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
               TabIndex        =   109
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
               TabIndex        =   105
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
               TabIndex        =   104
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
               TabIndex        =   65
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
               TabIndex        =   64
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
               TabIndex        =   63
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
               TabIndex        =   62
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
               TabIndex        =   57
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
               TabIndex        =   56
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
               TabIndex        =   55
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
               TabIndex        =   54
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
               TabIndex        =   53
               Top             =   390
               Width           =   1095
            End
            Begin VB.Label Label11 
               Caption         =   "Label7"
               Height          =   195
               Left            =   6510
               TabIndex        =   52
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
            TabIndex        =   99
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
            TabIndex        =   67
            Top             =   4620
            Width           =   1125
         End
         Begin VB.Label Label14 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   5880
            TabIndex        =   66
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
         TabIndex        =   36
         Top             =   660
         Width           =   5505
         Begin VB.TextBox aBusca 
            Height          =   285
            Index           =   3
            Left            =   3060
            MaxLength       =   8
            TabIndex        =   28
            Top             =   210
            Width           =   885
         End
         Begin VB.TextBox aBusca 
            Height          =   285
            Index           =   2
            Left            =   2400
            MaxLength       =   3
            TabIndex        =   27
            Top             =   210
            Width           =   435
         End
         Begin VB.CommandButton cBusca2 
            Caption         =   "&Buscar"
            Height          =   375
            Left            =   4260
            TabIndex        =   29
            Top             =   180
            Width           =   1095
         End
         Begin VB.ComboBox Combo3 
            Height          =   315
            Left            =   900
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Top             =   180
            Width           =   1425
         End
         Begin VB.Label Label21 
            Caption         =   "Tipo Doc."
            ForeColor       =   &H00800000&
            Height          =   225
            Left            =   90
            TabIndex        =   121
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
            TabIndex        =   38
            Top             =   240
            Width           =   195
         End
         Begin VB.Label Label9 
            Caption         =   "Tipo Doc."
            ForeColor       =   &H00800000&
            Height          =   105
            Left            =   30
            TabIndex        =   37
            Top             =   690
            Width           =   915
         End
      End
      Begin VB.Frame Fr2 
         Height          =   645
         Left            =   540
         TabIndex        =   9
         Top             =   420
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
         Height          =   6990
         Index           =   0
         Left            =   360
         TabIndex        =   7
         Top             =   1230
         Width           =   9435
         Begin VB.Frame Fr9 
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FF0000&
            Height          =   6495
            Left            =   30
            TabIndex        =   12
            Top             =   420
            Width           =   9240
            Begin VB.Frame FrameCaja 
               BackColor       =   &H0080FF80&
               Height          =   2415
               Left            =   240
               TabIndex        =   133
               Top             =   3600
               Width           =   8655
               Begin VB.CommandButton CmdAcepta 
                  Caption         =   "Aceptar"
                  Height          =   375
                  Left            =   7680
                  TabIndex        =   155
                  Top             =   1920
                  Width           =   855
               End
               Begin VB.CommandButton Cmdelimina 
                  Caption         =   "Elimina"
                  Height          =   375
                  Left            =   7560
                  TabIndex        =   154
                  Top             =   1200
                  Width           =   855
               End
               Begin VB.CommandButton CmdModifica 
                  Caption         =   "Modifica"
                  Height          =   375
                  Left            =   7560
                  TabIndex        =   153
                  Top             =   720
                  Width           =   855
               End
               Begin VB.CommandButton CmdAdiciona 
                  Caption         =   "Adiciona"
                  Height          =   375
                  Left            =   7560
                  TabIndex        =   152
                  Top             =   240
                  Width           =   855
               End
               Begin TextFer.TxFer TxFernumero 
                  Height          =   315
                  Left            =   4200
                  TabIndex        =   137
                  Top             =   1950
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
                  TabIndex        =   141
                  Top             =   1950
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
                  TabIndex        =   139
                  Top             =   1950
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
                  TabIndex        =   134
                  Top             =   1950
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
                  Left            =   2040
                  TabIndex        =   135
                  Top             =   1950
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
               Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
                  Height          =   1455
                  Left            =   120
                  TabIndex        =   151
                  Top             =   360
                  Width           =   6975
                  _ExtentX        =   12303
                  _ExtentY        =   2566
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
                  Left            =   6360
                  TabIndex        =   144
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
                  Left            =   4140
                  TabIndex        =   142
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
                  Left            =   5250
                  TabIndex        =   140
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
                  Left            =   2430
                  TabIndex        =   138
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
                  TabIndex        =   136
                  Top             =   120
                  Width           =   855
               End
            End
            Begin VB.TextBox TextContacto 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   3855
               MaxLength       =   100
               TabIndex        =   128
               Top             =   2205
               Width           =   4710
            End
            Begin VB.TextBox TxtEntrega 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1575
               MaxLength       =   70
               TabIndex        =   127
               Top             =   1260
               Width           =   6150
            End
            Begin MSComCtl2.DTPicker DTFecEnt 
               Height          =   315
               Left            =   4980
               TabIndex        =   123
               Top             =   840
               Width           =   1275
               _ExtentX        =   2249
               _ExtentY        =   556
               _Version        =   393216
               Format          =   17235969
               CurrentDate     =   39739
            End
            Begin VB.ComboBox Combo1 
               Height          =   315
               Left            =   3750
               Style           =   2  'Dropdown List
               TabIndex        =   20
               Top             =   3075
               Width           =   1995
            End
            Begin MSMask.MaskEdBox MBox 
               Height          =   285
               Index           =   0
               Left            =   1590
               TabIndex        =   13
               Top             =   120
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
               Left            =   2100
               TabIndex        =   14
               Top             =   120
               Width           =   585
               _ExtentX        =   1032
               _ExtentY        =   503
               _Version        =   393216
               Appearance      =   0
               Enabled         =   0   'False
               MaxLength       =   3
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MBox 
               Height          =   285
               Index           =   2
               Left            =   2760
               TabIndex        =   15
               Top             =   120
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   503
               _Version        =   393216
               Appearance      =   0
               Enabled         =   0   'False
               MaxLength       =   8
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MBox 
               Height          =   285
               Index           =   3
               Left            =   5340
               TabIndex        =   16
               Top             =   90
               Width           =   1215
               _ExtentX        =   2143
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
               Top             =   480
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
               Top             =   870
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
               Top             =   1695
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
               Top             =   2145
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
               Left            =   1560
               TabIndex        =   102
               Top             =   3045
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
               TabIndex        =   103
               Top             =   2655
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
               Left            =   7800
               TabIndex        =   124
               Top             =   855
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
               Left            =   1560
               TabIndex        =   130
               Top             =   2640
               Width           =   3795
               _ExtentX        =   6694
               _ExtentY        =   556
               XcodMaxLongitud =   3
               xcodwith        =   200
               NomTabla        =   "vt_modoventa"
               TituloAyuda     =   "Ayuda de Modo de ventas"
               ListaCampos     =   "modovtacodigo(1),modovtadescripcion(1),modovtaactctacte(3)"
               XcodCampo       =   "modovtacodigo"
               XListCampo      =   "modovtadescripcion"
               ListaCamposDescrip=   "Codigo,Descripcion, Cta Cte"
               ListaCamposText =   "modovtacodigo,modovtadescripcion,modovtaactctacte"
            End
            Begin MSMask.MaskEdBox MBox 
               Height          =   285
               Index           =   30
               Left            =   8040
               TabIndex        =   147
               Top             =   2640
               Width           =   555
               _ExtentX        =   979
               _ExtentY        =   503
               _Version        =   393216
               Appearance      =   0
               Enabled         =   0   'False
               MaxLength       =   2
               PromptChar      =   "_"
            End
            Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyuAlmacen 
               Height          =   315
               Left            =   6570
               TabIndex        =   149
               Top             =   3000
               Width           =   2505
               _ExtentX        =   4419
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
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Almacen"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Index           =   36
               Left            =   5760
               TabIndex        =   150
               Top             =   3060
               Width           =   795
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Moneda"
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
               Index           =   35
               Left            =   7080
               TabIndex        =   148
               Top             =   2640
               Width           =   705
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Nro.Pedido"
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
               Height          =   255
               Index           =   34
               Left            =   6600
               TabIndex        =   146
               Top             =   120
               Width           =   975
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
               TabIndex        =   132
               Top             =   6120
               Width           =   1125
            End
            Begin VB.Label Label8 
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
               ForeColor       =   &H000000C0&
               Height          =   315
               Left            =   6150
               TabIndex        =   131
               Top             =   6120
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
               Left            =   2745
               TabIndex        =   129
               Top             =   2250
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
               TabIndex        =   126
               Top             =   1305
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
               Left            =   7080
               TabIndex        =   125
               Top             =   885
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
               Left            =   3390
               TabIndex        =   122
               Top             =   900
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
               TabIndex        =   101
               Top             =   2715
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
               TabIndex        =   100
               Top             =   3105
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
               Height          =   255
               Index           =   0
               Left            =   240
               TabIndex        =   34
               Top             =   120
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
               TabIndex        =   33
               Top             =   540
               Width           =   615
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "F. Doc."
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
               Height          =   255
               Index           =   2
               Left            =   4440
               TabIndex        =   32
               Top             =   120
               Width           =   735
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
               TabIndex        =   31
               Top             =   1695
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
               TabIndex        =   30
               Top             =   2175
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
               Left            =   2220
               TabIndex        =   25
               Top             =   3135
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
               TabIndex        =   24
               Top             =   930
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
               TabIndex        =   23
               Top             =   480
               Width           =   4845
            End
            Begin VB.Label Label4 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000A&
               BorderStyle     =   1  'Fixed Single
               Caption         =   " "
               ForeColor       =   &H80000008&
               Height          =   285
               Left            =   2190
               TabIndex        =   22
               Top             =   1695
               Width           =   5535
            End
            Begin VB.Label Label7 
               BackColor       =   &H8000000E&
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
               Height          =   360
               Left            =   7620
               TabIndex        =   35
               Top             =   135
               Width           =   1485
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
            Y1              =   6975
            Y2              =   6975
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
            Top             =   -15
            Width           =   6075
         End
      End
      Begin VB.Frame Frame4 
         Height          =   930
         Index           =   0
         Left            =   3960
         TabIndex        =   6
         Top             =   8265
         Width           =   2100
         Begin VB.CommandButton cmdBotones 
            Caption         =   "&Acepta"
            Height          =   690
            Index           =   11
            Left            =   90
            Picture         =   "FrmCorrecciondatosgen.frx":1E22
            Style           =   1  'Graphical
            TabIndex        =   143
            Top             =   180
            Width           =   870
         End
         Begin VB.CommandButton cmdBotones 
            Caption         =   "&Cancelar"
            Height          =   690
            Index           =   12
            Left            =   1080
            Picture         =   "FrmCorrecciondatosgen.frx":2264
            Style           =   1  'Graphical
            TabIndex        =   145
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
      Top             =   9495
      Width           =   10575
      _ExtentX        =   18653
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
Attribute VB_Name = "Frmcorrecciondatosgen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adll As New dllgeneral.dll_general
Dim adicionactacte As Integer
Dim rs As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim rspago As New ADODB.Recordset
Dim almacen As String


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
    Dim nsql As String
    Set rspago = Nothing
    
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
           Set rs2 = VGCNx.Execute("select * from vt_vendedor where vendedorcodigo='" & MBox(6) & "'")
           If rs2.RecordCount > 0 Then
               Label4 = Escadena(rs2!vendedornombres)
           Else
             Label4 = ""
           End If
           rs2.Close
           MBox(7) = Escadena(rs!puntovtacodigo)
           MBox(24) = Escadena(rs!formapagocodigo)
           MBox(25) = Escadena(rs!modovtacodigo)
           MBox(30) = ESNULO(rs!pedidomoneda, "01")
           Ctr_Ayumodovta.xclave = MBox(25): Ctr_Ayumodovta.Ejecutar
           If IsNull(rs!pedidocondicionfactura) Then
               Combo1.ListIndex = 0
           Else
              Combo1.ListIndex = VerificaCombo(Combo1, rs!pedidocondicionfactura)
           End If
           Label8 = DatoMoneda(rs!pedidomoneda) & numero(rs!pedidototneto)
           TxtHor.Text = rs!horaentrega
           TextContacto.Text = rs!pedidoobserva
           almacen = Escadena(rs!almacencodigo)
           Ctr_AyuAlmacen.xclave = Escadena(rs!almacencodigo)
           If VGParamSistem.tesoreriaenlinea = 1 And VgModificar = 1 Or VgModificar = 0 And Format(rs!fechaact, "dd/mm/yyyy") = Format(Now, "dd/mm/yyyy") Then
              SQL = " select * from vt_pagosencaja where empresacodigo='" & VGParametros.empresacodigo & "'"
              SQL = SQL & " and pedidonumero='" & Label7 & "'"
              rs2.Open SQL, VGCNx, adOpenDynamic, adLockOptimistic
              If rs2.RecordCount > 0 Then
                 Set TDBGrid1.DataSource = rs2
                 FrameCaja.Visible = True
              End If
'              MBox(1).SetFocus
           End If
'           MBox(1).SetFocus
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
           almacen = Escadena(rs!almacencodigo)
           Ctr_AyuAlmacen.xclave = Escadena(rs2!almacencodigo)
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
           almacen = Escadena(rs!almacencodigo)
           Ctr_AyuAlmacen.xclave = Escadena(rs!almacencodigo)
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

Private Sub CmdAcepta_Click()
rs2!pagocodigo = Ctr_Ayuoperacion.xclave
rs2!pagotipocodigo = Ctr_Ayutipo.xclave
rs2!monedacodigo = TxFermoneda.valor
rs2!pagoimporte = TxFerimporte.valor
rs2!empresacodigo = VGParametros.empresacodigo
rs2!pedidonumero = Label7
rs2!pagonumdoc = TxFernumero.valor
rs2.UpdateBatch adAffectAllChapters

TDBGrid1.Refresh
End Sub

Private Sub CmdAdiciona_Click()
rs2.AddNew
rs2!pagocodigo = Ctr_Ayuoperacion.xclave
rs2!pagotipocodigo = Ctr_Ayutipo.xclave
rs2!monedacodigo = TxFermoneda.valor
rs2!pagoimporte = numero(TxFerimporte.valor)
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
                .Parameters("@almacen") = Ctr_AyuAlmacen.xclave
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
       Dim rsmov As New ADODB.Recordset
       Set rsmov = VGCNx.Execute(" select * from movalmcab where empresacodigo+canroped='" & VGParametros.empresacodigo & almacen & "'")
       If rsmov.RecordCount > 0 Then
          If almacen <> Ctr_AyuAlmacen.xclave Then
             MsgBox (" El pedido tiene movimientos de almacen , no se puede modificar ")
          End If
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
                .Parameters("@almacen") = Ctr_AyuAlmacen.xclave
                
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
           If VgModificar = 1 Then
              If VGParamSistem.tesoreriaenlinea Then actualizatesoreria
           End If
            aBusca(0) = ""
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
   Exit Sub
nerror:
   If Err Then
      MsgBox Err.Number & "-" & Err.Description
      MsgBox "Inconsistencia de Datos Referenciales ó Datos Incompletos", vbInformation, MsgTitle
      Err = 0
      Exit Sub
      Resume
   End If
End Sub

Private Sub actualizatesoreria()
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim VGCommandoSP  As New ADODB.Command
Dim adll As dllgeneral.dll_general
 SQL = " select * from te_cabecerarecibos where empresacodigo='" & VGParametros.empresacodigo & "'"
SQL = SQL & " and cabcomprobnumero='" & Label7 & "'and cabrec_estadoreg<>1 "
Set rs = VGCNx.Execute(SQL)
If rs.RecordCount > 0 Then
   rs.MoveFirst
   Do While Not rs.EOF
      If ESNULO(rs!comprobconta, 0) > 0 Then
         SQL = " delete ct_cabcomprob" & VGParamSistem.AnoProceso & " where empresacodigo='" & VGParametros.empresacodigo & "'"
         SQL = SQL & " and cabcomprobmes = " & Month(MBox(3)) & " and cabcomprobnumero='" & rs!comprobconta & "'"
         Set rs1 = VGCNx.Execute(SQL)
      End If
      SQL = "update te_cabecerarecibos set cabrec_estadoreg=1 where empresacodigo='" & VGParametros.empresacodigo & "'"
      SQL = SQL & " and cabrec_numrecibo='" & rs!cabrec_numrecibo & "'"
      Set rs1 = VGCNx.Execute(SQL)
      SQL = "update te_detallerecibos set detrec_estadoreg=1 where cabrec_numrecibo='" & rs!cabrec_numrecibo & "'"
      Set rs1 = VGCNx.Execute(SQL)
      rs.MoveNext
   Loop
End If
' SQL = "delete vt_pagosencaja where empresacodigo='" & VGParametros.empresacodigo & "'"
' SQL = SQL & " and pedidonumero='" & Label7 & "'"
' Set rs1 = VGCNx.Execute(SQL)
' Set rs1 = Nothing
'   SQL = " select top 0 * from vt_pagosencaja "
'   rs1.Open SQL, VGCNx, adOpenDynamic, adLockBatchOptimistic
'   rs1.AddNew
 '  rs1!empresacodigo = VGParametros.empresacodigo
  ' rs1!pedidonumero = Label7
'   rs1!pagocodigo = Ctr_Ayuoperacion.xclave
'   rs1!pagotipocodigo = Ctr_Ayutipo.xclave
'   rs1!pagonumdoc = TxFernumero.valor
'   rs1!monedacodigo = TxFermoneda.Text
'  rs1!cajerocodigo = "01"
'  rs1!pagoimporte = Format(Trim(TxFerimporte.valor), "###,###,##0.00")
'  rs1.UpdateBatch adAffectAllChapters
'   rs1.Close
  'Elimar los Detalle antes de grabar
  VGCommandoSP.ActiveConnection = VGgeneral
  VGCommandoSP.CommandType = adCmdStoredProc
  VGCommandoSP.CommandText = "vt_formadepago_pro"
  VGCommandoSP.Parameters.Refresh
  With VGCommandoSP
     .Parameters("@base") = VGParamSistem.BDEmpresa
     .Parameters("@empresa") = VGParametros.empresacodigo
     .Parameters("@pedido") = Label7
     Set rs1 = .Execute
  End With
  If rs1.RecordCount() > 0 Then
     rs1.MoveFirst
     Do While Not rs1.EOF
        Call grabatesoria(rs1)
        rs1.MoveNext
    Loop
  End If
 End Sub
Private Sub grabatesoria(ByVal rs As Recordset)
Dim Text1 As String
Dim acmd As New ADODB.Command
Dim rb As New ADODB.Recordset
Dim ingresacargo As Integer
Dim xabono, xzona, xmone, xcuenta, xtipo As String
Dim xnumplan, ximpsol, xtcam, xnumpag As Double
On Error GoTo error1
xtcam = XRecuperaTipoCambio(MBox(3), Venta, VGCNx)

VGCNx.BeginTrans
    'Actualizamos el numerador de tipo de ingreso
Set rb = VGCNx.Execute("select * from te_parametroempresa ")
    If rb.RecordCount > 0 Then
         Text1 = Right("0000000000" & Trim(CStr(CDbl(IIf(IsNull(rb!empresanumeingreso + 1) Or Len(Trim(rb!empresanumeingreso)) = 0, 1, rb!empresanumeingreso + 1)))), 6)
         VGCNx.Execute "Update te_parametroempresa Set empresanumeingreso='" & Right("0000000000" & Trim(CStr(Val(Text1))), 6) & "' "
     End If
rb.Close
Set rb = Nothing
VGCNx.CommitTrans
VGCNx.BeginTrans
    Set acmd.ActiveConnection = VGgeneral
    acmd.CommandType = adCmdStoredProc
    acmd.CommandText = "te_abonadocumento_pro"
    acmd.CommandTimeout = 0
    acmd.Prepared = True
    With acmd
        .Parameters("@base") = VGCNx.DefaultDatabase
        .Parameters("@tipo") = "1"
        .Parameters("@numrecibo") = Escadena(Text1)
        .Parameters("@estadoreg") = ""
        .Parameters("@controlctacte") = "1"
        .Parameters("@vendedorcodigo") = MBox(6)
        .Parameters("@cajacodigo") = RTrim(rs!banco)
        .Parameters("@clientecodigo") = MBox(4)
        .Parameters("@descripcion") = ""
        If rs.Fields(0) = "C" Then
           .Parameters("@operacion") = "03"
        Else
           .Parameters("@operacion") = "04"
        End If
        .Parameters("@monedacodigo") = rs!moneda
        .Parameters("@ingsal") = "I"
        .Parameters("@tipocambio") = xtcam
        .Parameters("@totsoles") = Round(IIf(MBox(30) = "01", rs!importe, rs!importe * xtcam), 2)
        .Parameters("@totdolares") = Round(IIf(MBox(30) = "02", rs!importe, rs!importe / xtcam), 2)
        .Parameters("@fechadocumento") = MBox(3)
        .Parameters("@empresa") = VGParametros.empresacodigo
        .Parameters("@observa") = ""
        .Parameters("@transferauto") = ""
        .Parameters("@numreciboegreso") = ""
        .Parameters("@usuario") = g_usuario
        .Parameters("@cabprovinumero") = Label7
        .Parameters("@fechaact") = Now
     End With
     acmd.Execute
     Set acmd = Nothing
       xmone = MBox(30)
             Set rb = VGCNx.Execute("select * from cc_tipodocumento where tdocumentocodigo='" & rs.Fields(1) & "'")
             xzona = "01": xnumpag = 1
             If rb.RecordCount > 0 Then
                xabono = rb!tdocumentotipo
                xtipo = IIf(IsNull(rb!tdocumentopermiteaplica), Null, rb!tdocumentopermiteaplica)
                If rs.Fields(7) = g_TipoSol Then
                   xcuenta = "" & Trim(rb!tdocumentocuentasoles)
                Else
                   xcuenta = "" & Trim(rb!tdocumentocuentadolares)
                End If
             Else
                xabono = "": xcuenta = "": xtipo = ""
             End If
             rb.Close
             Set rb = Nothing
        
             ' Registramos datos en Tesoreria
             Set acmd.ActiveConnection = VGgeneral
             acmd.CommandType = adCmdStoredProc
             acmd.CommandText = "te_abonadetalledocumento_pro"
             acmd.CommandTimeout = 0
             acmd.Prepared = True
             With acmd
                 .Parameters("@base") = VGCNx.DefaultDatabase
                 .Parameters("@tipo") = "1"
                 .Parameters("@numrecibo") = Text1
                 .Parameters("@estadoreg") = ""
                 .Parameters("@item") = "1"
                 .Parameters("@emisioncheque") = rs.Fields(0)
                 .Parameters("@tipodocconcepto") = rs.Fields(1)
                 .Parameters("@numdocumento") = rs.Fields(2)
                 .Parameters("@carabo") = xabono
                 .Parameters("@formacan") = rs.Fields(3)
                 .Parameters("@tdqc") = rs.Fields(4)
                 .Parameters("@ndqc") = Trim(rs.Fields(6))
                 .Parameters("@tipocajabanco") = rs.Fields(0)
                 .Parameters("@cajabanco") = RTrim(rs!banco)
                 .Parameters("@numctacte") = Escadena(rs.Fields(10))    'numero de cuenta corriente con tamaño 30
                 .Parameters("@adicionactacte") = "C"
                 .Parameters("@monedadocumento") = xmone
                 .Parameters("@monedacancela") = Escadena(rs.Fields(7))
                 .Parameters("@importesoles") = CDbl(IIf(rs.Fields(7) = g_TipoSol, rs.Fields(8), (rs.Fields(8) * xtcam)))
                 .Parameters("@importedolares") = CDbl(IIf(rs.Fields(7) = g_TipoSol, (rs.Fields(8) / xtcam), rs.Fields(8)))
                 .Parameters("@contabledisponi") = 0      'sale de empresas
                 .Parameters("@fechacancela") = rs.Fields(9)
                 .Parameters("@observacion") = Escadena(rs.Fields(11))
                 .Parameters("@cliente") = MBox(4)
                 .Parameters("@usuario") = g_usuario
                 .Parameters("@fechaact") = Now
             End With
             acmd.Execute
             Set acmd = Nothing
             DoEvents
    
    
VGCNx.CommitTrans
Call GeneraAsientoEnlineaTesor(MBox(3), VGParametros.empresacodigo, rs!tipo, Escadena(Text1), 1, "''''", MBox(30), "C", "E")
 MsgBox "Los datos han sido grabados satisfactoriamente...!!!", vbInformation, MsgTitle
 Exit Sub
error1:
  MsgBox "No se pudo Grabar " & Err.Description & " - " & Err.Number, vbInformation, Caption
  VGCNx.RollbackTrans

  Exit Sub
  Resume
 End Sub
           
Private Sub Cmdelimina_Click()
rs2.Delete
TDBGrid1.Refresh
End Sub

Private Sub CmdModifica_Click()
Ctr_Ayuoperacion.xclave = rs2!pagocodigo: Ctr_Ayuoperacion.Ejecutar
Ctr_Ayutipo.xclave = rs2!pagotipocodigo: Ctr_Ayutipo.Ejecutar
TxFermoneda.valor = rs2!monedacodigo: TxFermoneda.SetFocus
TxFerimporte.valor = rs2!pagoimporte: TxFerimporte.SetFocus
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

Private Sub Command1_Click()
rs2.AddNew
Ctr_Ayuoperacion.xclave = rs2!pagocodigo: Ctr_Ayuoperacion.Ejecutar
Ctr_Ayutipo.xclave = rs2!pagotipocodigo: Ctr_Ayutipo.Ejecutar
TxFermoneda.valor = rs2!monedacodigo: TxFermoneda.SetFocus
TxFerimporte.valor = rs2!pagoimporte: TxFerimporte.SetFocus
End Sub

Private Sub Command2_Click()

End Sub

Private Sub Command4_Click()
rs2!pagocodigo = Ctr_Ayuoperacion.xclave
rs2!pagotipocodigo = Ctr_Ayutipo.xclave
rs2!monedacodigo = TxFermoneda.valor
rs2!pagoimporte = TxFerimporte.valor
rs2.Update
End Sub

Private Sub Ctr_AyuAlmacen_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
Dim rsmov As New ADODB.Recordset
Set rsmov = VGCNx.Execute(" select * from movalmcab where empresacodigo+canroped='" & VGParametros.empresacodigo & almacen & "'")
If rsmov.RecordCount > 0 Then
   If rsmov!CAALMA <> Ctr_AyuAlmacen.xclave Then
      MsgBox (" El pedido tiene movimientos de almacen , no se puede modificar ")
   End If
End If
End Sub

Private Sub Ctr_Ayumodovta_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
If VgModificar = 1 Then
   adicionactacte = ColecCampos("modovtaactctacte")
Else
   adicionactacte = 1
End If
End Sub

Private Sub Ctr_Ayuoperacion_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
Ctr_Ayutipo.Filtro = "pagocodigo='" & Ctr_Ayuoperacion.xclave & "'"
Ctr_Ayutipo.xclave = ""
TxFernumero.valor = ""

If ColecCampos("pagoefectivo") = False Then
   Ctr_Ayutipo.Visible = False
   TxFernumero.Visible = False
Else
   Ctr_Ayutipo.Visible = True
   TxFernumero.Visible = True
End If
End Sub

Private Sub Form_Load()
   MostrarForm Me, "C"
   Fr9.Visible = False
   Fr10.Visible = False
   Fr11.Visible = False
   If VgModificar = 1 Then
      MBox(0).Enabled = True
      MBox(1).Enabled = True
      MBox(2).Enabled = True
   Else
      MBox(0).Enabled = False
      MBox(1).Enabled = False
      MBox(2).Enabled = False
   End If
   Call Ctr_Ayumodovta.conexion(VGCNx)
   Call Ctr_Ayuoperacion.conexion(VGCNx)
   Call Ctr_Ayutipo.conexion(VGCNx)
   Call Ctr_AyuAlmacen.conexion(VGCNx)
   Ctr_AyuAlmacen.Filtro = " empresacodigo='" & VGParametros.empresacodigo & "'"
   If VgModificar = 1 Then
        Ctr_Ayumodovta.Visible = True
        MBox(24).Visible = True
        Combo1.Visible = True
        FrameCaja.Visible = True
      Else
        Ctr_Ayumodovta.Visible = False
        MBox(24).Visible = False
        Combo1.Visible = False
        FrameCaja.Visible = False
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

