VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form FrmMovimientoProve 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Control de Cuentas por Pagar"
   ClientHeight    =   9645
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10275
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9645
   ScaleWidth      =   10275
   ShowInTaskbar   =   0   'False
   Begin VB.Frame FrameRetencion 
      Height          =   2055
      Left            =   720
      TabIndex        =   68
      Top             =   4140
      Visible         =   0   'False
      Width           =   8655
      Begin VB.CommandButton cayuda 
         Caption         =   "..."
         Height          =   315
         Index           =   11
         Left            =   2160
         TabIndex        =   95
         Top             =   720
         Width           =   255
      End
      Begin VB.TextBox text2 
         Height          =   285
         Index           =   16
         Left            =   6240
         MaxLength       =   11
         TabIndex        =   86
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox text2 
         Height          =   285
         Index           =   15
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   89
         Top             =   1170
         Width           =   5865
      End
      Begin VB.TextBox text2 
         Height          =   285
         Index           =   14
         Left            =   3240
         MaxLength       =   30
         TabIndex        =   75
         Top             =   360
         Width           =   2715
      End
      Begin VB.CommandButton cayuda 
         Caption         =   "..."
         Height          =   285
         Index           =   10
         Left            =   5970
         TabIndex        =   76
         Top             =   375
         Width           =   195
      End
      Begin VB.TextBox text2 
         Height          =   285
         Index           =   13
         Left            =   2490
         MaxLength       =   2
         TabIndex        =   73
         Top             =   360
         Width           =   465
      End
      Begin VB.TextBox text2 
         Height          =   285
         Index           =   12
         Left            =   6255
         MaxLength       =   11
         TabIndex        =   77
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton cayuda 
         Caption         =   "..."
         Height          =   315
         Index           =   9
         Left            =   2145
         TabIndex        =   72
         Top             =   360
         Width           =   255
      End
      Begin VB.CommandButton cayuda 
         Caption         =   "..."
         Height          =   315
         Index           =   8
         Left            =   2970
         TabIndex        =   74
         Top             =   360
         Width           =   255
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   7560
         TabIndex        =   93
         Top             =   1560
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   7560
         TabIndex        =   91
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox TxtMontoretencion 
         Height          =   285
         Left            =   7680
         TabIndex        =   88
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox TxtMontopagar 
         Height          =   285
         Left            =   7680
         TabIndex        =   83
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox text2 
         Height          =   285
         Index           =   18
         Left            =   1680
         MaxLength       =   2
         TabIndex        =   85
         Top             =   720
         Width           =   435
      End
      Begin VB.TextBox text2 
         Height          =   285
         Index           =   17
         Left            =   1680
         MaxLength       =   2
         TabIndex        =   71
         Top             =   360
         Width           =   435
      End
      Begin VB.TextBox TxtMonto 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
         EndProperty
         Height          =   405
         Left            =   120
         MaxLength       =   10
         TabIndex        =   70
         Top             =   1320
         Width           =   915
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Doc.Retencion"
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
         Height          =   180
         Index           =   21
         Left            =   120
         TabIndex        =   94
         Top             =   720
         Width           =   1365
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Doc.Cancelacion"
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
         Height          =   180
         Index           =   15
         Left            =   120
         TabIndex        =   92
         Top             =   360
         Width           =   1485
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Observaciones"
         Height          =   180
         Index           =   20
         Left            =   4110
         TabIndex        =   90
         Top             =   840
         Width           =   1155
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Nro. Cuenta Corriente"
         Height          =   180
         Index           =   19
         Left            =   3300
         TabIndex        =   87
         Top             =   120
         Width           =   2625
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "T.Canc."
         Height          =   180
         Index           =   18
         Left            =   1440
         TabIndex        =   84
         Top             =   120
         Width           =   645
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Banco"
         Height          =   180
         Index           =   17
         Left            =   2370
         TabIndex        =   82
         Top             =   120
         Width           =   615
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Numero"
         Height          =   180
         Index           =   16
         Left            =   6405
         TabIndex        =   81
         Top             =   120
         Width           =   885
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
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
         Index           =   14
         Left            =   0
         TabIndex        =   79
         Top             =   0
         Width           =   45
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Monto a Pagar"
         Height          =   300
         Index           =   13
         Left            =   7440
         TabIndex        =   78
         Top             =   120
         Width           =   1125
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Monto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   12
         Left            =   240
         TabIndex        =   69
         Top             =   1080
         Width           =   645
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00C0C0C0&
      Height          =   990
      Left            =   180
      TabIndex        =   42
      Top             =   8190
      Width           =   4230
      Begin VB.CommandButton cmdBotones 
         Caption         =   "&Salir"
         Height          =   690
         Index           =   7
         Left            =   3255
         Picture         =   "FrmMovimientoProve.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   57
         Top             =   210
         Width           =   870
      End
      Begin VB.CommandButton cmdBotones 
         Caption         =   "&Eliminar"
         Height          =   690
         Index           =   6
         Left            =   2250
         Picture         =   "FrmMovimientoProve.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   55
         Top             =   210
         Width           =   825
      End
      Begin VB.CommandButton cmdBotones 
         Cancel          =   -1  'True
         Caption         =   "&Grabar"
         Height          =   690
         Index           =   5
         Left            =   1260
         Picture         =   "FrmMovimientoProve.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   240
         Width           =   870
      End
      Begin VB.CommandButton cmdBotones 
         Caption         =   "&Nuevo"
         Height          =   690
         Index           =   4
         Left            =   180
         Picture         =   "FrmMovimientoProve.frx":0CC6
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   210
         Width           =   870
      End
   End
   Begin VB.Frame Frame1 
      Height          =   8025
      Left            =   135
      TabIndex        =   22
      Top             =   165
      Width           =   9975
      Begin VB.Frame Frame2 
         Height          =   2520
         Left            =   0
         TabIndex        =   31
         Top             =   0
         Width           =   9645
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
            Height          =   288
            Index           =   4
            Left            =   6300
            MaxLength       =   50
            TabIndex        =   6
            Top             =   960
            Width           =   3288
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H0080C0FF&
            BorderStyle     =   0  'None
            Height          =   765
            Left            =   5640
            TabIndex        =   58
            Top             =   1575
            Width           =   3915
            Begin VB.Label Label5 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               BorderStyle     =   1  'Fixed Single
               Enabled         =   0   'False
               Height          =   285
               Index           =   1
               Left            =   2430
               TabIndex        =   63
               Top             =   420
               Width           =   1365
            End
            Begin VB.Label Label5 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               BorderStyle     =   1  'Fixed Single
               Enabled         =   0   'False
               Height          =   285
               Index           =   0
               Left            =   480
               TabIndex        =   62
               Top             =   420
               Width           =   1365
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
               TabIndex        =   61
               Top             =   480
               Width           =   435
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
               TabIndex        =   60
               Top             =   480
               Width           =   345
            End
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
               TabIndex        =   59
               Top             =   0
               Width           =   3915
            End
         End
         Begin VB.CommandButton cayuda 
            Caption         =   "..."
            Height          =   315
            Index           =   0
            Left            =   1620
            TabIndex        =   11
            Top             =   600
            Width           =   225
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
            Left            =   1260
            MaxLength       =   2
            TabIndex        =   3
            Top             =   600
            Width           =   345
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   3
            Left            =   4650
            MaxLength       =   10
            TabIndex        =   10
            Top             =   2115
            Width           =   855
         End
         Begin VB.ComboBox Combo2 
            Height          =   315
            Left            =   1260
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   2085
            Width           =   1845
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
            TabIndex        =   0
            Top             =   210
            Width           =   1125
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
            Left            =   8280
            TabIndex        =   2
            Top             =   180
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
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
            Left            =   1260
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   180
            Width           =   2115
         End
         Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuda2 
            Height          =   345
            Left            =   5460
            TabIndex        =   4
            Top             =   600
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   609
            XcodMaxLongitud =   11
            xcodwith        =   800
            NomTabla        =   "cp_proveedor"
            TituloAyuda     =   "Ayuda de Proveedores"
            ListaCampos     =   "clientecodigo(1),clienterazonsocial(1),proveedorcontribuyente(2)"
            XcodCampo       =   "clientecodigo"
            XListCampo      =   "clienterazonsocial"
            ListaCamposDescrip=   "Codigo,Descripcion"
            ListaCamposText =   "clientecodigo,clienterazonsocial,proveedorcontribuyente"
         End
         Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyudaCaja 
            Height          =   315
            Left            =   1260
            TabIndex        =   5
            Top             =   945
            Width           =   3780
            _ExtentX        =   6668
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
         Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayutransf 
            Height          =   315
            Left            =   1260
            TabIndex        =   8
            Top             =   1650
            Visible         =   0   'False
            Width           =   4215
            _ExtentX        =   7435
            _ExtentY        =   556
            XcodMaxLongitud =   7
            xcodwith        =   800
            NomTabla        =   "te_cabecerarecibos"
            TituloAyuda     =   "Busqueda de Documentos x rendir"
            ListaCampos     =   "cabrec_numreciboegreso(1),cabrec_descripcion(1),SaldoDocxRendir(1),cabrec_fechadocumento(2),clientecodigo(1)"
            XcodCampo       =   "cabrec_numreciboegreso"
            XListCampo      =   "cabrec_descripcion"
            ListaCamposDescrip=   "Nro.transferencia,descripcion,Saldo,Fecha documento, cliente"
            ListaCamposText =   "cabrec_numreciboegreso,cabrec_descripcion,SaldoDocxRendir,cabrec_fechadocumento,clientecodigo"
         End
         Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuempresa 
            Height          =   315
            Left            =   1260
            TabIndex        =   7
            Top             =   1305
            Width           =   4230
            _ExtentX        =   7461
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
            Left            =   330
            TabIndex        =   98
            Top             =   1320
            Width           =   705
         End
         Begin VB.Label LeReferencia 
            AutoSize        =   -1  'True
            Caption         =   "Nro.Transf."
            Height          =   195
            Left            =   330
            TabIndex        =   97
            Top             =   1680
            Visible         =   0   'False
            Width           =   795
         End
         Begin VB.Label Label1 
            Caption         =   "Girado a:"
            Height          =   372
            Index           =   8
            Left            =   5520
            TabIndex        =   80
            Top             =   960
            Width           =   708
         End
         Begin VB.Label Label2 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Index           =   0
            Left            =   1920
            TabIndex        =   56
            Top             =   600
            Width           =   2565
         End
         Begin VB.Label Label1 
            Caption         =   "Moneda"
            Height          =   255
            Index           =   7
            Left            =   330
            TabIndex        =   40
            Top             =   2070
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "Tipo Cambio"
            Height          =   255
            Index           =   6
            Left            =   3690
            TabIndex        =   38
            Top             =   2145
            Width           =   1005
         End
         Begin VB.Label Label1 
            Caption         =   "Cod. Caja"
            Height          =   255
            Index           =   5
            Left            =   330
            TabIndex        =   37
            Top             =   960
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Proveedor"
            Height          =   255
            Index           =   4
            Left            =   4650
            TabIndex        =   36
            Top             =   630
            Width           =   825
         End
         Begin VB.Label Label1 
            Caption         =   "Operacion"
            Height          =   255
            Index           =   3
            Left            =   330
            TabIndex        =   35
            Top             =   600
            Width           =   885
         End
         Begin VB.Label Label1 
            Caption         =   "Ingreso/Egreso"
            Height          =   255
            Index           =   2
            Left            =   7050
            TabIndex        =   34
            Top             =   240
            Width           =   1155
         End
         Begin VB.Label Label1 
            Caption         =   "Recibo"
            Height          =   255
            Index           =   1
            Left            =   4470
            TabIndex        =   33
            Top             =   210
            Width           =   825
         End
         Begin VB.Label Label1 
            Caption         =   "Ingreso/Egreso"
            Height          =   255
            Index           =   0
            Left            =   60
            TabIndex        =   32
            Top             =   210
            Width           =   1455
         End
      End
      Begin VB.Frame Frame3 
         Height          =   5250
         Left            =   180
         TabIndex        =   39
         Top             =   2670
         Width           =   9675
         Begin VB.Frame Frame4 
            Height          =   1635
            Left            =   90
            TabIndex        =   41
            Top             =   3570
            Width           =   9465
            Begin VB.TextBox text2 
               Height          =   375
               Index           =   11
               Left            =   1320
               MaxLength       =   4
               TabIndex        =   100
               Top             =   360
               Width           =   495
            End
            Begin VB.TextBox text2 
               Height          =   375
               Index           =   2
               Left            =   1800
               MaxLength       =   10
               TabIndex        =   99
               Top             =   360
               Width           =   1095
            End
            Begin VB.TextBox text2 
               Height          =   285
               Index           =   10
               Left            =   3210
               MaxLength       =   50
               TabIndex        =   29
               Top             =   1320
               Width           =   6105
            End
            Begin VB.CommandButton cayuda 
               Caption         =   "..."
               Height          =   285
               Index           =   7
               Left            =   2850
               TabIndex        =   23
               Top             =   1305
               Width           =   195
            End
            Begin VB.TextBox text2 
               Height          =   285
               Index           =   9
               Left            =   120
               MaxLength       =   30
               TabIndex        =   21
               Top             =   1290
               Width           =   2715
            End
            Begin VB.CommandButton cayuda 
               Caption         =   "..."
               Height          =   285
               Index           =   6
               Left            =   6930
               TabIndex        =   26
               Top             =   390
               Width           =   195
            End
            Begin VB.CommandButton cayuda 
               Caption         =   "..."
               Height          =   315
               Index           =   5
               Left            =   4815
               TabIndex        =   19
               Top             =   390
               Width           =   255
            End
            Begin VB.CommandButton cayuda 
               Caption         =   "..."
               Height          =   315
               Index           =   4
               Left            =   3990
               TabIndex        =   17
               Top             =   390
               Width           =   255
            End
            Begin VB.CommandButton cayuda 
               Caption         =   "..."
               Height          =   315
               Index           =   3
               Left            =   2865
               TabIndex        =   14
               Top             =   390
               Width           =   255
            End
            Begin VB.CommandButton cayuda 
               Caption         =   "..."
               Height          =   315
               Index           =   2
               Left            =   1080
               TabIndex        =   13
               Top             =   390
               Width           =   255
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
               Height          =   300
               Left            =   8145
               TabIndex        =   28
               Top             =   390
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   10
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin VB.TextBox text2 
               Height          =   285
               Index           =   8
               Left            =   7170
               MaxLength       =   10
               TabIndex        =   27
               Top             =   390
               Width           =   915
            End
            Begin VB.TextBox text2 
               Height          =   285
               Index           =   6
               Left            =   5100
               MaxLength       =   11
               TabIndex        =   24
               Top             =   360
               Width           =   1335
            End
            Begin VB.TextBox text2 
               Height          =   285
               Index           =   5
               Left            =   4320
               MaxLength       =   2
               TabIndex        =   18
               Top             =   390
               Width           =   465
            End
            Begin VB.TextBox text2 
               Height          =   285
               Index           =   4
               Left            =   3525
               MaxLength       =   2
               TabIndex        =   16
               Top             =   390
               Width           =   435
            End
            Begin VB.TextBox text2 
               Height          =   285
               Index           =   3
               Left            =   3165
               MaxLength       =   1
               TabIndex        =   15
               Top             =   390
               Width           =   285
            End
            Begin VB.TextBox text2 
               Height          =   285
               Index           =   1
               Left            =   690
               MaxLength       =   2
               TabIndex        =   12
               Top             =   360
               Width           =   375
            End
            Begin VB.TextBox text2 
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
            Begin VB.TextBox text2 
               Height          =   285
               Index           =   7
               Left            =   6450
               MaxLength       =   2
               TabIndex        =   25
               Top             =   390
               Width           =   465
            End
            Begin VB.Label lblMonProv 
               BorderStyle     =   1  'Fixed Single
               Height          =   270
               Left            =   165
               TabIndex        =   67
               Top             =   705
               Width           =   885
            End
            Begin VB.Label Label6 
               BorderStyle     =   1  'Fixed Single
               Height          =   345
               Left            =   3060
               TabIndex        =   66
               Top             =   750
               Width           =   5850
            End
            Begin VB.Label Label3 
               Alignment       =   2  'Center
               Caption         =   "Observaciones"
               Height          =   180
               Index           =   11
               Left            =   3240
               TabIndex        =   65
               Top             =   1080
               Width           =   6075
            End
            Begin VB.Label Label3 
               Alignment       =   2  'Center
               Caption         =   "Nro. Cuenta Corriente"
               Height          =   180
               Index           =   10
               Left            =   180
               TabIndex        =   64
               Top             =   1050
               Width           =   2625
            End
            Begin VB.Label Label3 
               Alignment       =   2  'Center
               Caption         =   "Fec. Cancela"
               Height          =   210
               Index           =   9
               Left            =   8280
               TabIndex        =   52
               Top             =   150
               Width           =   975
            End
            Begin VB.Label Label3 
               Alignment       =   2  'Center
               Caption         =   "Importe"
               Height          =   210
               Index           =   8
               Left            =   7260
               TabIndex        =   51
               Top             =   150
               Width           =   765
            End
            Begin VB.Label Label3 
               Alignment       =   2  'Center
               Caption         =   "Mon."
               Height          =   180
               Index           =   7
               Left            =   6570
               TabIndex        =   50
               Top             =   150
               Width           =   465
            End
            Begin VB.Label Label3 
               Alignment       =   2  'Center
               Caption         =   "Numero"
               Height          =   180
               Index           =   6
               Left            =   5190
               TabIndex        =   49
               Top             =   150
               Width           =   1245
            End
            Begin VB.Label Label3 
               Alignment       =   2  'Center
               Caption         =   "Banco"
               Height          =   180
               Index           =   5
               Left            =   4275
               TabIndex        =   48
               Top             =   150
               Width           =   615
            End
            Begin VB.Label Label3 
               Alignment       =   2  'Center
               Caption         =   "T.Canc."
               Height          =   180
               Index           =   4
               Left            =   3465
               TabIndex        =   47
               Top             =   150
               Width           =   645
            End
            Begin VB.Label Label3 
               Alignment       =   2  'Center
               Caption         =   "T/P"
               Height          =   180
               Index           =   3
               Left            =   3060
               TabIndex        =   46
               Top             =   150
               Width           =   465
            End
            Begin VB.Label Label3 
               Alignment       =   2  'Center
               Caption         =   "Numero"
               Height          =   180
               Index           =   2
               Left            =   1590
               TabIndex        =   45
               Top             =   150
               Width           =   1065
            End
            Begin VB.Label Label3 
               Alignment       =   2  'Center
               Caption         =   "Tipo"
               Height          =   180
               Index           =   1
               Left            =   690
               TabIndex        =   44
               Top             =   150
               Width           =   405
            End
            Begin VB.Label Label3 
               Alignment       =   2  'Center
               Caption         =   "Item"
               Height          =   180
               Index           =   0
               Left            =   90
               TabIndex        =   43
               Top             =   150
               Width           =   645
            End
         End
         Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
            Height          =   3375
            Left            =   120
            TabIndex        =   96
            Top             =   240
            Width           =   9375
            _ExtentX        =   16536
            _ExtentY        =   5953
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
            Columns(1).Caption=   "tipodoc"
            Columns(1).DataField=   "detrec_tipodoc_concepto"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "numero doc"
            Columns(2).DataField=   "detrec_numdocumento"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "T/P"
            Columns(3).DataField=   "detrec_forcan"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "tdqc"
            Columns(4).DataField=   "tdqc"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Banco"
            Columns(5).DataField=   "detrec_cajabanco1"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "ndqc"
            Columns(6).DataField=   "detrec_ndqc"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   0
            Columns(7)._MaxComboItems=   5
            Columns(7).Caption=   "monedadocumento"
            Columns(7).DataField=   "detrec_monedadocumento"
            Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(8)._VlistStyle=   0
            Columns(8)._MaxComboItems=   5
            Columns(8).Caption=   "Importe soles"
            Columns(8).DataField=   "importe"
            Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(9)._VlistStyle=   0
            Columns(9)._MaxComboItems=   5
            Columns(9).Caption=   "fecha cancelacion"
            Columns(9).DataField=   "detrec_fechacancela"
            Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(10)._VlistStyle=   0
            Columns(10)._MaxComboItems=   5
            Columns(10).Caption=   "Importe dolares"
            Columns(10).DataField=   "cargo"
            Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(11)._VlistStyle=   0
            Columns(11)._MaxComboItems=   5
            Columns(11).Caption=   "Codigo Emp"
            Columns(11).DataField=   "empresacodigo"
            Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   12
            Splits(0)._UserFlags=   0
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   13160660
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=12"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=714"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=635"
            Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(5)=   "Column(1).Width=609"
            Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=529"
            Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(9)=   "Column(2).Width=1931"
            Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=1852"
            Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(13)=   "Column(3).Width=635"
            Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=556"
            Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(17)=   "Column(4).Width=714"
            Splits(0)._ColumnProps(18)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(19)=   "Column(4)._WidthInPix=635"
            Splits(0)._ColumnProps(20)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(21)=   "Column(5).Width=953"
            Splits(0)._ColumnProps(22)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(23)=   "Column(5)._WidthInPix=873"
            Splits(0)._ColumnProps(24)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(25)=   "Column(6).Width=1508"
            Splits(0)._ColumnProps(26)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(27)=   "Column(6)._WidthInPix=1429"
            Splits(0)._ColumnProps(28)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(29)=   "Column(7).Width=688"
            Splits(0)._ColumnProps(30)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(31)=   "Column(7)._WidthInPix=609"
            Splits(0)._ColumnProps(32)=   "Column(7).Order=8"
            Splits(0)._ColumnProps(33)=   "Column(8).Width=1931"
            Splits(0)._ColumnProps(34)=   "Column(8).DividerColor=0"
            Splits(0)._ColumnProps(35)=   "Column(8)._WidthInPix=1852"
            Splits(0)._ColumnProps(36)=   "Column(8).Order=9"
            Splits(0)._ColumnProps(37)=   "Column(9).Width=1879"
            Splits(0)._ColumnProps(38)=   "Column(9).DividerColor=0"
            Splits(0)._ColumnProps(39)=   "Column(9)._WidthInPix=1799"
            Splits(0)._ColumnProps(40)=   "Column(9).Order=10"
            Splits(0)._ColumnProps(41)=   "Column(10).Width=2223"
            Splits(0)._ColumnProps(42)=   "Column(10).DividerColor=0"
            Splits(0)._ColumnProps(43)=   "Column(10)._WidthInPix=2143"
            Splits(0)._ColumnProps(44)=   "Column(10).Order=11"
            Splits(0)._ColumnProps(45)=   "Column(11).Width=1693"
            Splits(0)._ColumnProps(46)=   "Column(11).DividerColor=0"
            Splits(0)._ColumnProps(47)=   "Column(11)._WidthInPix=1614"
            Splits(0)._ColumnProps(48)=   "Column(11).Order=12"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   3
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
            AllowAddNew     =   -1  'True
            AllowUpdate     =   0   'False
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
            _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=32,.bold=0,.fontsize=825,.italic=0"
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
            _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=66,.parent=13"
            _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=63,.parent=14"
            _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=64,.parent=15"
            _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=65,.parent=17"
            _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=70,.parent=13"
            _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=67,.parent=14"
            _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=68,.parent=15"
            _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=69,.parent=17"
            _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=74,.parent=13"
            _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=71,.parent=14"
            _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=72,.parent=15"
            _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=73,.parent=17"
            _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=78,.parent=13"
            _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=75,.parent=14"
            _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=76,.parent=15"
            _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=77,.parent=17"
            _StyleDefs(80)  =   "Splits(0).Columns(11).Style:id=82,.parent=13"
            _StyleDefs(81)  =   "Splits(0).Columns(11).HeadingStyle:id=79,.parent=14"
            _StyleDefs(82)  =   "Splits(0).Columns(11).FooterStyle:id=80,.parent=15"
            _StyleDefs(83)  =   "Splits(0).Columns(11).EditorStyle:id=81,.parent=17"
            _StyleDefs(84)  =   "Named:id=33:Normal"
            _StyleDefs(85)  =   ":id=33,.parent=0"
            _StyleDefs(86)  =   "Named:id=34:Heading"
            _StyleDefs(87)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(88)  =   ":id=34,.wraptext=-1"
            _StyleDefs(89)  =   "Named:id=35:Footing"
            _StyleDefs(90)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(91)  =   "Named:id=36:Selected"
            _StyleDefs(92)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(93)  =   "Named:id=37:Caption"
            _StyleDefs(94)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(95)  =   "Named:id=38:HighlightRow"
            _StyleDefs(96)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(97)  =   "Named:id=39:EvenRow"
            _StyleDefs(98)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(99)  =   "Named:id=40:OddRow"
            _StyleDefs(100) =   ":id=40,.parent=33"
            _StyleDefs(101) =   "Named:id=41:RecordSelector"
            _StyleDefs(102) =   ":id=41,.parent=34"
            _StyleDefs(103) =   "Named:id=42:FilterBar"
            _StyleDefs(104) =   ":id=42,.parent=33"
         End
      End
   End
   Begin MSComctlLib.StatusBar Panel 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   20
      Top             =   9300
      Width           =   10275
      _ExtentX        =   18124
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
Attribute VB_Name = "FrmMovimientoProve"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adll As New dllgeneral.dll_general
Dim contribuyente As String
Dim Tabla As String
Dim adicionaretenciones As Integer
Dim Emiteretencion As Integer
Dim rsdetat As New ADODB.Recordset
Dim monto As Double
Dim fecharendicion As Date
Dim controlarendicion As Boolean
Dim tipooperacion As Integer
Dim xempresacodigo As String
Dim m_fondofijo As Integer
Dim m_docxrendir As Integer
Dim fechatransferencia As Date
Dim saldodocxrendir As Double
Dim clientecodigo As String

Property Let docxrendir(valor As String)
   m_docxrendir = valor
End Property
Property Let fondofijo(valor As String)
   m_fondofijo = valor
End Property

Public Sub cargar_grilla()
   Tabla = VGcomputer + "det_recibo"
   Set rsdetat = Nothing
   If ExisteElem(0, VGCNx, Tabla) Then VGCNx.Execute ("drop table " & Tabla)
   
   SQL = " select top 0 item=detrec_item,detrec_tipodoc_concepto,detrec_numdocumento,detrec_forcan, "
   SQL = SQL & " tdqc=detrec_tdqc,detrec_cajabanco1,detrec_ndqc,detrec_monedadocumento,importe=detrec_importesoles, "
   SQL = SQL & " detrec_fechacancela,detrec_numctacte,detrec_observacion,retencion='0',empresacodigo='01',cargo=detrec_importedolares,"
   SQL = SQL & " detrec_carabo into " & Tabla & " from te_detallerecibos "
   Set rsdetat = VGCNx.Execute(SQL)
   SQL = " select * from " & Tabla
 '  Call rsdetat.Fields.Append("Item", adChar, 3)
 '  Call rsdetat.Fields.Append("Tipo", adChar, 2)
 '  Call rsdetat.Fields.Append("Numero", adChar, 11)
 '  Call rsdetat.Fields.Append("T/P", adChar, 1)
 '  Call rsdetat.Fields.Append("T.Canc", adChar, 2)
 '  Call rsdetat.Fields.Append("Banco", adChar, 2)
 '  Call rsdetat.Fields.Append("Numero Doc", adChar, 20)
 '  Call rsdetat.Fields.Append("Mnda", adChar, 2)
 '  Call rsdetat.Fields.Append("Importe", adDouble)
 '  Call rsdetat.Fields.Append("Fecha Canc", adDate)
 '  Call rsdetat.Fields.Append("Cta Cte", adChar, 30)
 '  Call rsdetat.Fields.Append("Observaciones", adChar, 50)
 '  Call rsdetat.Fields.Append("Retencion", adChar, 1)
 '  Call rsdetat.Fields.Append("empresacodigo", adChar, 2)
 '  Call rsdetat.Fields.Append("cargo", adDouble)
 '  Call rsdetat.Fields.Append("carabo", adChar, 1)
   rsdetat.Open SQL, VGCNx, adOpenDynamic, adLockBatchOptimistic
   Set TDBGrid1.DataSource = rsdetat
   TDBGrid1.Refresh
   
End Sub

Private Sub cAyuda_Click(Index As Integer)
 Dim rb As New ADODB.Recordset
 Dim nMonedaCab As String
 Dim SQL As String
 nAyuda = "": nDetalle = ""
 nAyuda1 = "": nMoneda = ""
 
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
         FrmAyudaTes.BCondi = "operacioncontrolaclienteprov='" & IIf(adll.ComboDato(Combo1) = "I", "P", "P") & "'"
         FrmAyudaTes.BFiltro = dfiltra
         FrmAyudaTes.Show 1
         Text1(1) = nAyuda
         Label2(0) = nDetalle
         'ctr_ayudacaja.xclave.SetFocus
         Call Text1_KeyPress(1, 13)
         
    ElseIf Index = 1 Then
         Set rb = VGCNx.Execute("select * from te_operaciongeneral where operacioncodigo='" & Escadena(Text1(1)) & "' and operacioncontrolaclienteprov='" & IIf(adll.ComboDato(Combo1) = "I", "P", "P") & "'")
         If rb.RecordCount > 0 Then
            If Escadena(rb!operacionvalidacajabancos) = "B" Then
                Ctr_AyudaCaja.Visible = False
                cayuda(1).Enabled = False
                Combo2.SetFocus
                rb.Close
                Set rb = Nothing
                Exit Sub
            Else
                Ctr_AyudaCaja.Visible = True
                cayuda(1).Enabled = True
                Ctr_AyudaCaja.SetFocus
            End If
        End If
        rb.Close
        Set rb = Nothing
         
        If Len(Trim(Ctr_AyudaCaja.xclave)) > 0 Then
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
        Ctr_AyudaCaja.xclave = nAyuda
        Label2(1) = nDetalle
        SendKeys "{tab}"
     ElseIf Index = 2 Then
       If Len(Trim(Text2(1))) > 0 Then
          SendKeys "{tab}"
          Exit Sub
        End If
        SQL = "select * from cp_tipodocumento where tdocumentoingplan='1' "
        If adll.ComboDato(Combo1) = "I" Then
           SQL = SQL & " and  tdocumentoactualizaxtesoreria='1'"
         End If
         If adll.VerificaDatoExistente(VGCNx, SQL) = 1 Then
            Dim zfiltra(1, 2) As String
            SQL = " tdocumentoingplan='1' "
           If adll.ComboDato(Combo1) = "I" Then
              SQL = SQL & " and  tdocumentoactualizaxtesoreria='1'"
            End If
            zfiltra(1, 1) = "Documento": zfiltra(1, 2) = "tdocumentocodigo"
            FrmAyudaTes.TipoForma = 1
            FrmAyudaTes.BConexion = VGCNx
            FrmAyudaTes.Bdata = "0"
            FrmAyudaTes.BTabla = "cp_tipodocumento"
            FrmAyudaTes.BCampos = "tdocumentocodigo as Codigo,tdocumentodescripcion as Descripcion"
            FrmAyudaTes.BOrden = "tdocumentocodigo"
            FrmAyudaTes.BCondi = SQL
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
         If adll.VerificaDatoExistente(VGCNx, "select * from cp_cargo where empresacodigo='" & Ctr_Ayuempresa.xclave & "' and clientecodigo='" & Ctr_Ayuda2.xclave & "' and documentocargo='" & Text2(1).Text & "'") = 1 Then
            Dim wfiltra(1, 2) As String
            wfiltra(1, 1) = "Documento": wfiltra(1, 2) = "cargonumdoc"
            FrmAyudaTes.TipoForma = 5
            FrmAyudaTes.BConexion = VGCNx
            FrmAyudaTes.Bdata = "0"
            FrmAyudaTes.BTabla = "cp_cargo A inner join cp_tipodocumento B On a.documentocargo=b.tdocumentocodigo"
            FrmAyudaTes.BCampos = "documentocargo as TD,cargonumdoc as Numero,monedacodigo as Mnd,cargoapeimpape as Total,(Round(cargoapeimpape,2)-Round(isnull(cargoapeimppag,0),2)) as Saldo,cargoapefecemi as FecEmision, cargoapefecvct as FecVencimiento, cargoaperefere as Referencia"
            FrmAyudaTes.BOrden = "Clientecodigo,cargoapefecemi"
            If adll.ComboDato(Combo1) = "E" Then
               FrmAyudaTes.BCondi = "empresacodigo ='" & Ctr_Ayuempresa.xclave & "' and clientecodigo='" & Trim(Ctr_Ayuda2.xclave) & "' and isnull(cargoapeflgcan,0)<>1 and b.tdocumentotipo='C' and a.documentocargo='" & Trim(Text2(1).Text) & "' and isnull(cargoapeflgreg,0)<>1"
            Else
               FrmAyudaTes.BCondi = "empresacodigo ='" & Ctr_Ayuempresa.xclave & "' and clientecodigo='" & Trim(Ctr_Ayuda2.xclave) & "' and isnull(cargoapeflgcan,0)<>1 and b.tdocumentotipo='A' and a.documentocargo='" & Trim(Text2(1).Text) & "' and isnull(cargoapeflgreg,0)<>1"
            End If
            FrmAyudaTes.BFiltro = gfiltra
            FrmAyudaTes.Show 1
            If nAyuda = Empty Then Exit Sub
            If Len(nDetalle) > 11 Then
                Text2(11).Text = Left(nDetalle, 4)
                Text2(2).Text = Mid(nDetalle, 5, Len(nDetalle) - 4)
            Else
                Text2(11).Text = Left(nDetalle, 3)
                Text2(2).Text = Mid(nDetalle, 4, Len(nDetalle) - 3)
            End If
            Text2(8).Text = nAyuda
            Label6.Caption = nAyuda1
            lblMonProv.Caption = nMoneda
            nMonedaCab = Left(Combo2.List(Combo2.ListIndex), 2)
            Text2(7).Text = nMonedaCab
            If lblMonProv.Caption <> nMonedaCab Then
               If nMonedaCab = g_tiposol Then
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
  ElseIf Index = 4 Or Index = 9 Or Index = 11 Then   'Tipo de cancelacion
    If Len(Trim(Text2(4))) > 0 And Index = 4 Then
      SendKeys "{tab}"
      Exit Sub
    End If
    If Len(Trim(Text2(17))) > 0 And Index = 9 Then
      SendKeys "{tab}"
      Exit Sub
    End If
    If Len(Trim(Text2(18))) > 0 And Index = 11 Then
      SendKeys "{tab}"
      Exit Sub
    End If
    If adll.VerificaDatoExistente(VGCNx, "select * from cp_tipodocumento where tdocumentotipo='A' and tdocumentocancela='1'") = 1 Then
        Dim ffiltra(1, 2) As String
        ffiltra(1, 1) = "Documento": ffiltra(1, 2) = "tdocumentocodigo"
        FrmAyudaTes.TipoForma = 1
        FrmAyudaTes.BConexion = VGCNx
        FrmAyudaTes.Bdata = "0"
        FrmAyudaTes.BTabla = "cp_tipodocumento"
        FrmAyudaTes.BCampos = "tdocumentocodigo as Codigo,tdocumentodescripcion as Descripcion"
        FrmAyudaTes.BOrden = "tdocumentocodigo"
        FrmAyudaTes.BCondi = "tdocumentotipo='A' and tdocumentocancela='1'"
        FrmAyudaTes.BFiltro = ffiltra
        FrmAyudaTes.Show 1
        Text2(4).Text = nAyuda
        Text2(17).Text = nAyuda
        If Index = 4 Then Call Text2_KeyPress(4, 13)
        If Index = 9 Then Call Text2_KeyPress(17, 13)
     Else
         nAyuda = "": nDetalle = ""
         MsgBox "No existen Documentos...", vbInformation, MsgTitle
         Exit Sub
     End If
     Exit Sub
   ElseIf Index = 5 Or Index = 8 Then    'Tipo de Banco
        If Len(Trim(Text2(5))) > 0 And Index = 5 Then
          SendKeys "{tab}"
          Exit Sub
        End If
        If Len(Trim(Text2(13))) > 0 And Index = 8 Then
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
        Text2(13) = nAyuda
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
   ElseIf Index = 7 Or Index = 10 Then   'Nro Cuenta Corriente
        If Len(Trim(Text2(9))) > 0 And Index = 7 Then
           SendKeys "{tab}"
           Exit Sub
        End If
        If Len(Trim(Text2(14))) > 0 And Index = 14 Then
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
        Text2(14) = nAyuda
  ElseIf Index = 11 Then   'Tipo de retencion
    If Len(Trim(Text2(18))) > 0 Then
      SendKeys "{tab}"
      Exit Sub
    End If
    SQL = "select * from cp_tipodocumento where tdocumentotipo='A' and tdocumentocancela='1' and tdocumentocodigo='"
    SQL = SQL & VGParametros.empresacodigoretencion & "'"
    If adll.VerificaDatoExistente(VGCNx, SQL) = 1 Then
        ffiltra(1, 1) = "Documento": ffiltra(1, 2) = "tdocumentocodigo"
        FrmAyudaTes.TipoForma = 1
        FrmAyudaTes.BConexion = VGCNx
        FrmAyudaTes.Bdata = "0"
        FrmAyudaTes.BTabla = "cp_tipodocumento"
        FrmAyudaTes.BCampos = "tdocumentocodigo as Codigo,tdocumentodescripcion as Descripcion"
        FrmAyudaTes.BOrden = "tdocumentocodigo"
        FrmAyudaTes.BCondi = "tdocumentotipo='A' and tdocumentocancela='1'"
        FrmAyudaTes.BFiltro = ffiltra
        FrmAyudaTes.Show 1
        Text2(18).Text = nAyuda
        If Index = 4 Then Call Text2_KeyPress(4, 13)
        If Index = 9 Then Call Text2_KeyPress(17, 13)
     Else
         nAyuda = "": nDetalle = ""
         MsgBox "No existen Documentos...", vbInformation, MsgTitle
         Exit Sub
     End If
     Exit Sub
   End If
   nAyuda = "": nDetalle = ""
End Sub

Public Function GrabarData() As Integer
  Dim acmd As New ADODB.Command
  Dim rb As New ADODB.Recordset
  Dim xabono, xzona, xmone, xcuenta, xtipo, nosaldos As String
  Dim xnumplan, ximpsol, xtcam, xnumpag As Double
  Dim rsql As String
  Dim grabauno, xactualizaxtesoreria As String
 On Error GoTo error
    GrabarData = 0
    grabauno = 0
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
    VGCNx.CommitTrans
    
    Set rb = Nothing
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
        .Parameters("@fechadocumento") = Format(MBox1.Text, "dd/mm/yyyy")
        .Parameters("@observa") = Escadena(Text1(4))
        .Parameters("@transferauto") = ""
        .Parameters("@numreciboegreso") = Ctr_Ayutransf.xclave
        If VGParametros.sistemamultiempresas Then
           .Parameters("@empresa") = Ctr_Ayuempresa.xclave
         Else
           .Parameters("@empresa") = VGParametros.empresacodigo
        End If
        .Parameters("@usuario") = VGUsuario
        .Parameters("@fechaact") = Now
        If Ctr_Ayutransf.Visible = True Then
           .Parameters("@NumeroDocXRendir") = Ctr_Ayutransf.xclave
           .Parameters("@responsablectasxrendir") = clientecodigo
        End If
        
     End With
     acmd.Execute
     Set acmd = Nothing
     xmone = adll.ComboDato(Combo2)

     If rsdetat.RecordCount > 0 Then
                  
        '**** Actualizamos correlativo de documentos de cancelacion
         rsdetat.MoveFirst
         Do Until rsdetat.EOF
            If grabauno = 0 Then
               If rsdetat.Fields(4) = VGParametros.empresacodigoretencion And RTrim(rsdetat.Fields(6)) <> "" Then grabauno = RTrim(rsdetat.Fields(6))
            End If
            rsdetat.MoveNext
         Loop
         rsdetat.MoveFirst
         rsql = " select tdocumentonumeauto,tdocumentonumerador from cp_tipodocumento "
         rsql = rsql & " where tdocumentocodigo='" & VGParametros.empresacodigoretencion & "'"
         Set rb = VGCNx.Execute(rsql)
         If rb.RecordCount() > 0 Then
            If Not rb!tdocumentonumeauto And grabauno = 0 Then
               grabauno = Format(rb!tdocumentonumerador, "00000000000")
               rsql = Format(grabauno + 1, "00000000000")
               VGCNx.Execute "Update  cp_tipodocumento Set tdocumentonumerador='" & rsql & "' where tdocumentocodigo='" & VGParametros.empresacodigoretencion & "'"
            End If
         End If
         Do Until rsdetat.EOF
            nosaldos = "0"
             Set rb = VGCNx.Execute("select * from cp_tipodocumento where tdocumentocodigo='" & rsdetat.Fields(1) & "'")
             If rb.RecordCount > 0 Then
                xabono = rb!tdocumentotipo
                xactualizaxtesoreria = IIf(IsNull(rb!tdocumentoactualizaxtesoreria), 0, rb!tdocumentoactualizaxtesoreria)
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
             Dim abono As Integer
             abono = 0
             Set rb = VGCNx.Execute("select * from cp_cargo where empresacodigo='" & Ctr_Ayuempresa.xclave & "' and clientecodigo='" & Ctr_Ayuda2.xclave & "' and  documentocargo='" & rsdetat.Fields(1) & "' and cargonumdoc='" & rsdetat.Fields(2) & "'")
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
             
             ximpsol = Round(CDbl(rsdetat.Fields(8)), 2)
             xtcam = CDbl(Text1(3).Text)
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
         If (xabono = "C" And adll.ComboDato(Combo1) = "E" Or adll.ComboDato(Combo1) = "I" And Not xabono = "C") Then
             If rsdetat.Fields(4) = VGParametros.empresacodigoretencion Then rsdetat.Fields(6) = grabauno
             If rsdetat.Fields(4) = VGParametros.empresacodigoretencion Then nosaldos = "1"
             Set acmd.ActiveConnection = VGGeneral
             acmd.CommandType = adCmdStoredProc
             acmd.CommandText = "cp_abonadocumento_pro"
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
                 .Parameters("@ctabanco") = Escadena(rsdetat.Fields(10))      'Cuenta Banco
                 .Parameters("@abonoflpres") = "1"
                 .Parameters("@abonocanimpcan") = CDbl(rsdetat.Fields(8))
                 .Parameters("@abonocanimpsol") = ximpsol
                 .Parameters("@usuario") = VGUsuario
                 .Parameters("@fechaact") = Now
                 .Parameters("@forma") = rsdetat.Fields(3)
                 .Parameters("@monedacan") = rsdetat.Fields(7)
                 .Parameters("@abonocantd") = rsdetat.Fields(4)
                 .Parameters("@abonocannro") = Trim(rsdetat.Fields(6))
                 .Parameters("@fechacan") = rsdetat.Fields(9)
                 .Parameters("@cliente") = Escadena(Ctr_Ayuda2.xclave)
                 If VGParametros.sistemamultiempresas Then
                     .Parameters("@empresa") = Ctr_Ayuempresa.xclave
                   Else
                     .Parameters("@empresa") = VGParametros.empresacodigo
                 End If
             End With
             acmd.Execute
             
             Set acmd = Nothing
             DoEvents
             '**** Actualizamos Saldos de documento pendiente
             If rsdetat.Fields(7) = g_tipodolar Then
                If xmone = g_tiposol Then
                    VGCNx.Execute "Update  cp_cargo Set cargoapeimppag=isnull(cargoapeimppag,0)+" & CDbl(rsdetat.Fields(8) * xtcam) & "," & _
                                " cargoapenumpag='" & xnumpag + 1 & "'" & _
                               " Where documentocargo='" & rsdetat.Fields(1) & "' and cargonumdoc='" & Trim(rsdetat.Fields(2)) & "' and " & _
                               " clientecodigo='" & Trim(Ctr_Ayuda2.xclave) & "'"
                Else
                    VGCNx.Execute "Update  cp_cargo Set cargoapeimppag=isnull(cargoapeimppag,0)+" & CDbl(rsdetat.Fields(8)) & "," & _
                               " cargoapenumpag='" & xnumpag + 1 & "'" & _
                               " Where documentocargo='" & rsdetat.Fields(1) & "' and cargonumdoc='" & Trim(rsdetat.Fields(2)) & "' and " & _
                               " clientecodigo='" & Trim(Ctr_Ayuda2.xclave) & "'"
                End If
             ElseIf rsdetat.Fields(7) = g_tiposol Then
                If xmone = g_tipodolar Then
                    VGCNx.Execute "Update  cp_cargo Set cargoapeimppag=isnull(cargoapeimppag,0)+" & CDbl(rsdetat.Fields(8) / xtcam) & "," & _
                               " cargoapenumpag='" & xnumpag + 1 & "'" & _
                               " Where documentocargo='" & rsdetat.Fields(1) & "' and cargonumdoc='" & Trim(rsdetat.Fields(2)) & "' and " & _
                               " clientecodigo='" & Trim(Ctr_Ayuda2.xclave) & "'"
                Else
                    VGCNx.Execute "Update  cp_cargo Set cargoapeimppag=isnull(cargoapeimppag,0)+" & CDbl(rsdetat.Fields(8)) & "," & _
                               " cargoapenumpag='" & xnumpag + 1 & "'" & _
                               " Where documentocargo='" & rsdetat.Fields(1) & "' and cargonumdoc='" & Trim(rsdetat.Fields(2)) & "' and " & _
                               " clientecodigo='" & Trim(Ctr_Ayuda2.xclave) & "'"
                End If
             End If
             
             VGCNx.Execute "Update  cp_cargo " & _
                         " Set cargoapeflgcan= CASE Round(isnull(cargoapeimpape,0),2)-Round(isnull(cargoapeimppag,0),2) WHEN 0 THEN '1' ELSE '0' END ," & _
                         "   cargoapefeccan='" & rsdetat.Fields(9) & "'" & _
                         " Where documentocargo='" & rsdetat.Fields(1) & "' and cargonumdoc='" & Trim(rsdetat.Fields(2)) & "' and " & _
                         " clientecodigo='" & Trim(Ctr_Ayuda2.xclave) & "'"
             
           Else
             '****Permite Anticipos
             ' Registramos datos en cp_cargos
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
                  
                  '**** Actualizamos correlativo de documentos de anticipos
                    
                   rsql = " select tdocumentonumeauto,tdocumentonumerador from cp_tipodocumento "
                   rsql = rsql & " where tdocumentocodigo='" & Trim(rsdetat.Fields(1)) & "'"
                   Set rb = VGCNx.Execute(rsql)
                   If rb!tdocumentonumeauto = 1 And rb!tdocumentonumerador = Trim(rsdetat.Fields(2)) Then
                      rsdetat.Fields(2) = rb!tdocumentonumerador
                      rsql = Format(rsdetat.Fields(2) + 1, "00000000000")
                      VGCNx.Execute "Update  cp_tipodocumento Set tdocumentonumerador='" & rsql & "' where tdocumentocodigo='" & Trim(rsdetat.Fields(1)) & "'"
                   End If
                  Set acmd.ActiveConnection = VGGeneral
                  acmd.CommandType = adCmdStoredProc
                  acmd.CommandText = "cp_ingresacargo_pro"
                  acmd.CommandTimeout = 0
                  acmd.Prepared = True
                         With acmd
                           .Parameters("@base") = VGCNx.DefaultDatabase
                           .Parameters("@tipo") = 1
                           .Parameters("@tabla") = "cp_cargo"
                            If VGParametros.sistemamultiempresas Then
                                .Parameters("@empresa") = Ctr_Ayuempresa.xclave
                            Else
                                .Parameters("@empresa") = VGParametros.empresacodigo
                            End If
                           .Parameters("@tipodocu") = rsdetat.Fields(1)
                           .Parameters("@numero") = Trim(rsdetat.Fields(2))
                           .Parameters("@cliente") = Escadena(Ctr_Ayuda2.xclave)
                           .Parameters("@abononumplanilla") = Right("00000000" & Trim(Text1(0)), 6)
                           .Parameters("@vendedor") = ""
                           .Parameters("@abonotipoplanilla") = "TE"
                           .Parameters("@zona") = xzona
                           .Parameters("@apefecemi") = rsdetat.Fields(9)
                           .Parameters("@moneda") = rsdetat.Fields(7)
                           .Parameters("@apeimppag") = CDbl(rsdetat.Fields(8))
                           .Parameters("@usuario") = VGUsuario
                           .Parameters("@tipocambio") = CDbl(xtcam)
                           .Parameters("@fechaact") = Now
                           .Parameters("@flagcancel") = 0
                           .Parameters("@cargoabono") = "C"
                           .Parameters("@concepto") = ""
                           .Parameters("@glosa") = " "
                         End With
                         acmd.Execute
                         
                         Set acmd = Nothing
                         DoEvents
                                         
     
End If

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
                 .Parameters("@emisioncheque") = IIf(Len(Trim(Ctr_AyudaCaja.xclave)) = 0, "B", "C") ' ver si es cheque
                 .Parameters("@tipodocconcepto") = rsdetat.Fields(1)
                 .Parameters("@numdocumento") = Trim(rsdetat.Fields(2))
                 .Parameters("@carabo") = xabono
                 .Parameters("@formacan") = rsdetat.Fields(3)
                 .Parameters("@tdqc") = rsdetat.Fields(4)
                 .Parameters("@ndqc") = Trim(rsdetat.Fields(6))
                 .Parameters("@tipocajabanco") = IIf(Len(Trim(Ctr_AyudaCaja.xclave)) = 0, "B", "C")
                 .Parameters("@cajabanco") = IIf(Len(Trim(Ctr_AyudaCaja.xclave)) = 0, Escadena(rsdetat.Fields(5)), Trim(Ctr_AyudaCaja.xclave))
                 .Parameters("@numctacte") = Escadena(rsdetat.Fields(10))    'numero de cuenta corriente
                 .Parameters("@adicionactacte") = "P"
                 .Parameters("@monedadocumento") = xmone
                 .Parameters("@monedacancela") = adll.ComboDato(Combo2)
                 .Parameters("@importesoles") = CDbl(IIf(rsdetat.Fields(7) = g_tiposol, rsdetat.Fields(8), (rsdetat.Fields(8) * xtcam)))
                 .Parameters("@importedolares") = CDbl(IIf(rsdetat.Fields(7) = g_tiposol, (rsdetat.Fields(8) / xtcam), rsdetat.Fields(8)))
                 .Parameters("@contabledisponi") = Escadena(VGParametros.saldocontadispo)      'sale de empresas
                 .Parameters("@fechacancela") = rsdetat.Fields(9)
                 .Parameters("@observacion") = Escadena(rsdetat.Fields(11))
                 .Parameters("@gastos") = ""
                 .Parameters("@usuario") = VGUsuario
                 .Parameters("@fechaact") = Now
                 .Parameters("@nosaldos") = nosaldos
                 .Parameters("@cliente") = Escadena(Ctr_Ayuda2.xclave)
             End With
             acmd.Execute
             Set acmd = Nothing
             DoEvents
             rsdetat.MoveNext
         Loop
    End If
    rsdetat.Close
    VGCNx.CommitTrans
    Set rsdetat = Nothing
    If VGParametros.controlaestadosrendicion Then
       If controlarendicion Then Call Actualizasaldorendicion
    End If
    GrabarData = 1
    MsgBox "Los datos han sido grabados satisfactoriamente...!!!", vbInformation, MsgTitle
Exit Function
error:
  VGCNx.RollbackTrans
  MsgBox "No se pudo Grabar " & Err.Description & " - " & Err.Number, vbInformation, Caption
Exit Function
Resume
 
End Function

Private Sub cmdBotones_Click(Index As Integer)
  Dim nvalor As String
  
 On Error Resume Next
  
  Select Case Index
   
    Case 4
       Frame4.Enabled = True
       Call Limpiartexto(Text2, 0, 8)
       Frame4.Enabled = False
       Frame2.Enabled = True
       Call Limpiartexto(Text1, 0, 3)
       Set rsdetat = Nothing
       Call cargar_grilla
       Combo1.SetFocus
       Label2(0).Caption = Empty
       Label2(1).Caption = Empty
       adicionaretenciones = 0
    
    Case 5
      If ValidarGrabacion() = 1 Then
         cmdBotones(5).Enabled = False
         'Grabamos Cabecera de Tesoreria
         If GrabarData() = 1 And adicionaretenciones = 0 Then
               
           'Generando Asiento Contable en Linea para cuentas por cobrar
           If VGParametros.sistemaasientoenlinea Then
              Call GeneraAsientoEnlineaTesor(CDate(MBox1.Text), Ctr_Ayuempresa.xclave, "X", Escadena(Text1(0)), 1, "''''", adll.ComboDato(Combo2), IIf(Len(Trim(Ctr_AyudaCaja.xclave)) = 0, "B", "C"), adll.ComboDato(Combo1))
           End If
           If MsgBox("Desea Imprimir el Recibo ", vbQuestion + vbOKCancel) = vbOK Then
             Call ImprimirRecibo(Escadena(Text1(0)))
           End If
           If VGParametros.empresaretencion = 1 And adll.ComboDato(Combo1) = "E" Then
              If MsgBox("Desea Imprimir el Comprobante de retencion ", vbQuestion + vbOKCancel) = vbOK Then
                 Call ImprimirComprobanteretencion(Escadena(Text1(0)))
              End If
           End If
         Else
           MsgBox "No se guardaron los datos....!!!", vbInformation, MsgTitle
          End If
         cmdBotones(5).Enabled = True
         Frame2.Enabled = True
         Call Limpiartexto(Text1, 0, 3)
         Combo1.SetFocus
         Call cmdBotones_Click(4)
         adicionaretenciones = 0
      End If
      If adicionaretenciones = 1 Then Call adicionaretencion
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
      
    Case 7
      Unload Me
  End Select
End Sub

Function ValidarGrabacion() As Integer
   Dim numdoc As Integer
   Dim monto As Double
   Dim montotot As Double
   Dim numretencion As Integer
   Dim valorretencion As Double
   Dim xmsg As String
   Dim xcarabo As String
   Dim rsaux As New ADODB.Recordset
   Dim xrendicion As String
   ValidarGrabacion = 0
   If rsdetat.RecordCount <= 0 Then
     MsgBox "Falta a�adir el Detalle a la Ventana del Browse", vbInformation, Caption
     ValidarGrabacion = 0
     Exit Function
   End If
   If VGParametros.sistemamultiempresas Then
      If Ctr_Ayuempresa.xclave = "" Then
         MsgBox "Debe ingresar codigo de empresa ", vbInformation
        Exit Function
      End If
   End If
   rsdetat.MoveFirst
   monto = 0
   montotot = 0
   numdoc = 0
   numretencion = 0
   valorretencion = 0
   xcarabo = rsdetat!detrec_carabo
   If adll.ComboDato(Combo1) = "E" Then
     Do Until rsdetat.EOF()
       If rsdetat!detrec_carabo <> "A" Then
         If rsdetat.Fields(4) <> VGParametros.empresacodigoretencion Then montotot = montotot + rsdetat.Fields(14)
         If rsdetat!retencion <> 2 Then
            monto = monto + rsdetat.Fields(8)
            If rsdetat.Fields(4) <> VGParametros.empresacodigoretencion Or rsdetat!retencion = 1 And rsdetat.Fields(4) <> VGParametros.empresacodigoretencion Then
               numdoc = numdoc + 1
             ElseIf rsdetat.Fields(4) = VGParametros.empresacodigoretencion Then
                    numretencion = numretencion + 1
                    valorretencion = valorretencion + rsdetat.Fields(8)
            End If
         End If
       Else
         If rsdetat!detrec_carabo <> xcarabo Then
            MsgBox "Los Pagos son en documentos diferentes a los anticipos ", vbInformation, Caption
             Exit Function
         End If
      End If
      rsdetat.MoveNext
    Loop
   End If
  If Ctr_Ayutransf.Visible = True And montotot > saldodocxrendir Then
       MsgBox "Monto del documento ingresado excede al saldo de doc. a rendir, que es  -- > " & saldodocxrendir & " , verifique ", vbInformation
      Ctr_Ayutransf.SetFocus
      Exit Function
    End If
   If xcarabo = "A" Or adll.ComboDato(Combo1) = "I" Then
      ValidarGrabacion = 1
      Exit Function
   End If
    If Not contribuyente = True Then
       If monto <= VGParametros.minimoretencion Then
          If numretencion > numdoc Then
             MsgBox "Existe Documentos de Retencion para Monto Minimpo de " & VGParametros.sistemaminimoretencion, vbInformation, Caption
             Exit Function
           End If
        ElseIf numdoc > numretencion Then
                  xmsg = "Falta a�adir " & numdoc - numretencion & " Comprobantes de Retencion, desea agregar retenciones  "
                  If MsgBox(xmsg, vbYesNo) = vbYes Then
                     adicionaretenciones = 1
                    Else
                      ValidarGrabacion = 1
                  End If
                  Exit Function
            ElseIf numdoc < numretencion Then
                  MsgBox "Falta a�adir " & numretencion - numdoc & " Documentos a Pagar ", vbInformation, Caption
                  Exit Function
       End If
       If montotot >= VGParametros.minimoretencion And adicionaretenciones = 0 Then
          MsgBox "Monto a Pagar es mayor a Monto Minimpo de Retencion " & VGParametros.minimoretencion, vbInformation, Caption
          If MsgBox(" Desea agregar retenciones ", vbYesNo) = vbYes Then
             adicionaretenciones = 1
           Else
             ValidarGrabacion = 1
          End If
          Exit Function
        Else
          If adicionaretenciones = 1 Then adicionaretenciones = 0
        End If
 End If
   
   If tipooperacion = 1 And VGParametros.sistemabancarizacion = True Then
      If Left(Combo2.Text, 2) = "01" Then
         MsgBox "Monto a Pagar : " & monto & "  en SOLES por Caja es mayor a  " & VGParametros.sistemabancarizacion01 & " Es Operacion BANCOS ", vbInformation, Caption
         Exit Function
      End If
      If Left(Combo2.Text, 2) = "02" Then
         MsgBox "Monto a Pagar : " & monto & " en DOLARES por Caja es mayor a  " & VGParametros.sistemabancarizacion02 & " Es Operacion BANCOS ", vbInformation, Caption
         Exit Function
      End If
   End If
   If VGParametros.controlaestadosrendicion Then
      If controlarendicion Then
          SQL = "select numero=max(rendicionnumero) from te_rendiciones where oficinacodigo='" & VGoficina & "'"
          SQL = SQL & " and codigocaja='" & Ctr_AyudaCaja.xclave & "'"
          Set rsaux = VGCNx.Execute(SQL)
          xrendicion = rsaux!numero
          SQL = " select rendicionfecha,rendicionsaldofinal=rendicionsaldoinicial+rendicioningresos-rendicionegresos + isnull(saldoacumuladoxrendir,0)"
          SQL = SQL & " from te_rendiciones where oficinacodigo='" & VGoficina & "'"
          SQL = SQL & " and codigocaja='" & Ctr_AyudaCaja.xclave & "' and rendicionnumero='" & xrendicion & "'"
          Set rsaux = VGCNx.Execute(SQL)
          If numero(rsaux!rendicionsaldofinal) - Label5(0) < 0 Then
             MsgBox " Monto de cancelacion Origina Saldos Negativos, Monto permitido es de  -- > " & rsaux!saldofinal, vbInformation
             Ctr_AyudaCaja.SetFocus
            Exit Function
          End If
       End If
    End If
   ValidarGrabacion = 1
End Function

Private Sub CmdSalir_Click()
 FrameRetencion.Visible = False
 SendKeys "{tab}"
End Sub

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

Private Sub Command1_Click()
 Text2(4).Text = Text2(17).Text
 Text2(8).Text = TxtMontopagar
 Text2(6).Text = Text2(12).Text
 If Len(Trim(Ctr_AyudaCaja.xclave)) <> 0 And Text2(8).Text > 0 Then
    Call MBox2_KeyDown(13, 0)
  Else
    Call Text2_KeyPress(10, 13)
 End If
 Text2(4).Text = Text2(18)
 Text2(8).Text = TxtMontoretencion
 Text2(6).Text = Text2(16).Text
 If Text2(8).Text > 0 Then
    If Len(Trim(Ctr_AyudaCaja.xclave)) <> 0 And Text2(8).Text > 0 Then
       Call MBox2_KeyDown(13, 0)
     Else
       Call Text2_KeyPress(10, 13)
    End If
 End If
 FrameRetencion.Visible = False
 Call Limpiartexto(Text2, 0, 8)
End Sub

Private Sub Ctr_Ayuda2_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
contribuyente = IIf(IsNull(ColecCampos("proveedorcontribuyente")), 0, ColecCampos("proveedorcontribuyente"))
End Sub

Private Sub Ctr_AyudaCaja_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
Dim rsaux As New ADODB.Recordset
Dim xrendicion As String
If ColecCampos!cajarendiciones Then
   SQL = "select numero=max(rendicionnumero) from te_rendiciones where oficinacodigo='" & VGoficina & "'"
   SQL = SQL & " and codigocaja='" & Ctr_AyudaCaja.xclave & "'"
   Set rsaux = VGCNx.Execute(SQL)
   If rsaux.RecordCount > 0 And ESNULO(rsaux!numero, 0) > 0 Then
      xrendicion = rsaux!numero
      SQL = " select rendicionfecha from te_rendiciones where oficinacodigo='" & VGoficina & "'"
     SQL = SQL & " and codigocaja='" & Ctr_AyudaCaja.xclave & "' and rendicionnumero='" & xrendicion & "'"
     Set rsaux = VGCNx.Execute(SQL)
     fecharendicion = rsaux!rendicionfecha - VGParametros.diasatrazorendicion
     controlarendicion = ColecCampos!cajarendiciones
   Else
     controlarendicion = False
   End If
End If
If m_docxrendir = 1 Or m_fondofijo = 1 Then
      LeReferencia.Visible = True
      Ctr_Ayutransf.Visible = True
      Ctr_Ayutransf.filtro = " isnull(estadodocxrendir,0)<2 and cajacodigo='" & Ctr_AyudaCaja.xclave & "' and cabrec_transferenciaautomatico=1 "
 Else
      LeReferencia.Visible = False
      Ctr_Ayutransf.Visible = False
End If
End Sub

Private Sub Ctr_Ayutransf_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
fechatransferencia = ColecCampos("cabrec_fechadocumento")
saldodocxrendir = ColecCampos("saldodocxrendir")
clientecodigo = ColecCampos("clientecodigo")
End Sub

Private Sub Form_Load()
   MostrarForm Me, "C"
   
   Combo1.Clear
   Combo1.AddItem "I- INGRESOS"
   Combo1.AddItem "E- EGRESOS"
   Combo1.ListIndex = 0
   
   Call Ctr_Ayuda2.conexion(VGCNx)
   Call Ctr_AyudaCaja.conexion(VGCNx)
   Ctr_AyudaCaja.filtro = " isnull(CajaCuentaxRendir,0)=" & m_docxrendir & " and isnull(Cajafondofijo,0)=" & m_fondofijo
   Call Ctr_Ayutransf.conexion(VGCNx)
   Call Ctr_Ayuempresa.conexion(VGCNx)
   If VGParametros.sistemamultiempresas Then
      Ctr_Ayuempresa.Visible = True
    Else
      Ctr_Ayuempresa.xclave = VGParametros.empresacodigo
      Ctr_Ayuempresa.Visible = False
      Lblempresa.Visible = False
   End If
   Text1(0).Enabled = False
   Call adll.llenacombo(Combo2, "select monedacodigo,monedadescripcion from gr_moneda where monedacodigo<>'00'", VGCNx)
   Combo2.ListIndex = 0
   adicionaretenciones = 0
   Frame4.Enabled = False
   tipooperacion = 0
   MBox1 = Format(VGParamSistem.fechatrabajo, "dd/mm/yyyy")
   Text1(3) = numero(DatoTipoCambio(VGcnxCT, MBox1))
   Call cargar_grilla
   
End Sub

Private Sub MBox1_KeyDown(KeyCode As Integer, Shift As Integer)
  Call Seguir(MBox1, KeyCode)
End Sub

Private Sub MBox1_LostFocus()
 If IsDate(MBox1.Text) Then Text1(3).Text = DatoTipoCambio(VGcnxCT, MBox1.Text)
End Sub

Private Sub MBox2_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lEncontro As Boolean
  lEncontro = False
  If KeyCode = 13 Then
     If m_docxrendir = 1 And Format(MBox2.Text, "dd/mm/yyyy") < fechatransferencia Then
        MsgBox "La Fecha de Cancelaci�n es mayor a fecha de entrega de mportes" & fechatransferencia, vbInformation, Caption

     End If
     If Format(MBox2.Text, "dd/mm/yyyy") <> Format(MBox1.Text, "dd/mm/yyyy") Then
        MsgBox "La Fecha de Cancelaci�n debe ser la misma para todos los Documentos", vbInformation, Caption
        MBox2.Text = Format(MBox1.Text, "dd/mm/yyyy")
        MBox2.SetFocus
        Exit Sub
     End If
     If Len(Trim(Ctr_AyudaCaja.xclave)) = 0 Then
         SendKeys "{tab}"
     Else
        Call grabacion
      End If
  End If
End Sub

Public Sub grabacion()
   Dim rb As New ADODB.Recordset
   Dim rb1 As New ADODB.Recordset
   Dim Emiteretencion As String
   Dim totdoc As Double
   Dim fechadoc As Date
   Dim carabo As String
   Dim SQL As String
   Set rb = VGCNx.Execute("select * from cp_tipodocumento where tdocumentoingplan='1' and tdocumentocodigo='" & Text2(1) & "'")
   If rb.RecordCount = 0 Then
      MsgBox "No existe tipo de documento...Verifique!!", vbInformation, MsgTitle
      rb.Close
      Set rb = Nothing
      Text2(1).SetFocus
      Exit Sub
    Else
      carabo = rb!tdocumentotipo
   End If
    rb.Close
    Set rb = Nothing
    SQL = "select * from cp_cargo where clientecodigo='" & Ctr_Ayuda2.xclave
    SQL = SQL & "' and documentocargo='" & Text2(1) & "' and cargonumdoc='" & Text2(11).Text & Text2(2).Text & "'"
    Set rb = VGCNx.Execute(SQL)
    If rb.RecordCount() > 0 Then
       Emiteretencion = IIf(IsNull(rb!cargoemiteretencion), 0, rb!cargoemiteretencion)
       totdoc = rb!cargoapeimpape
       fechadoc = rb!cargoapefecemi
       xempresacodigo = Escadena(rb!empresacodigo)
     Else
       fechadoc = MBox2.Text
       Emiteretencion = 0
       xempresacodigo = "00"
    End If
    rb.Close
    Set rb = Nothing
    If MBox2.Text < fechadoc Then   '***JCGI Aca cambiariamos el filtro para q permita anticipo planilla
       MsgBox "fecha de Cancelacion Menor a fecha de Documento del dia ---> " & fechadoc, vbInformation, MsgTitle
       MBox2.SetFocus
       Exit Sub
    End If
    If VGParametros.controlaestadosrendicion Then
       If MBox2.Text < fecharendicion Then
          MsgBox "fecha de Cancelacion Menor a fecha ---> " & fecharendicion & " , MINIMA de las rendiciones ", vbInformation, MsgTitle
          MBox2.SetFocus
          Exit Sub
       End If
    End If
    If Not Text2(3) Like "[TP]" Then
      MsgBox "Solo debe ingresar P � T", vbInformation, MsgTitle
      Text2(3).SetFocus
      Exit Sub
    End If
    If Text2(2) = "" Then
      MsgBox "Solo debe ingresar Valor en Numero de documento", vbInformation, MsgTitle
      Text2(2).SetFocus
      Exit Sub
    End If
    
    Set rb = VGCNx.Execute("select * from cp_tipodocumento where tdocumentotipo='A' and tdocumentoingcobra='1' and tdocumentocodigo='" & Text2(4) & "'")
    If rb.RecordCount = 0 Then
      MsgBox "No existe tipo de documento...Verifique!!", vbInformation, MsgTitle
      rb.Close
      Set rb = Nothing
      Text2(4).SetFocus
      Exit Sub
    End If
    If rb!tdocumentovalidabanco = 1 And Not Ctr_AyudaCaja.Visible And Text2(5) = "" Then
      MsgBox "No existe Codigo de banco...Verifique!!", vbInformation, MsgTitle
      rb.Close
      Set rb = Nothing
      cayuda(5).Enabled = True
      Text2(5).Enabled = True
      Text2(5).SetFocus
      Exit Sub
    End If
    If rb!tdocumentovalidabanco = 1 And Text2(5) <> "" And Text2(6) = "" Then
      MsgBox "No existe Nro de cheque u Operacion...Verifique!!", vbInformation, MsgTitle
      rb.Close
      Set rb = Nothing
      Text2(6).Enabled = True
      Text2(6).SetFocus
      Exit Sub
    End If
    If rb!tdocumentovalidabanco = 1 And Not Ctr_AyudaCaja.Visible And Text2(9) = "" Then
      MsgBox "No existe Nro de cuenta bancaria...Verifique!!", vbInformation, MsgTitle
      rb.Close
      Set rb = Nothing
      Text2(9).SetFocus
      Exit Sub
    End If
    rb.Close
    Set rb = Nothing
    '***JCGI
    If VerificaNumDocIngresado = True Then
        MsgBox "Documento ya ha sido ingresado, no puede volver a ingresarlo", vbExclamation
        Text2(2).SetFocus
        Exit Sub
    End If

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
    
    rsdetat.AddNew
    rsdetat.Fields(0) = Escadena(Text2(0))
    rsdetat.Fields(1) = Escadena(Text2(1))
    rsdetat.Fields(2) = Format(Escadena(Text2(11).Text), "0000") & Format(Escadena(Text2(2)), "0000000000")
    rsdetat.Fields(3) = Escadena(Text2(3))
    rsdetat.Fields(4) = Escadena(Text2(4))
    rsdetat.Fields(5) = Escadena(Text2(5))
    rsdetat.Fields(6) = Escadena(Text2(6))
    rsdetat.Fields(7) = Escadena(Text2(7))
    rsdetat.Fields(8) = numero(Text2(8))
    rsdetat.Fields(9) = Format(MBox2, "dd/mm/yyyy")
    rsdetat.Fields(10) = Escadena(Text2(9).Text)
    rsdetat.Fields(11) = Escadena(Text2(10).Text)
    rsdetat.Fields(12) = Emiteretencion
    rsdetat.Fields(14) = totdoc
    rsdetat.Fields(15) = carabo
    rsdetat!empresacodigo = xempresacodigo
    
    rsdetat.Update
    TDBGrid1.Refresh
    If VGParametros.empresaretencion = 0 Then
       Call Limpiartexto(Text2, 0, 8)
    End If
    Text2(0) = CStr(CDbl(rsdetat.Fields(0)) + 1)
    Call Totales
    Text2(1).SetFocus
End Sub

Private Sub MBox2_LostFocus()
Call MBox2_KeyDown(13, 0)
End Sub
Private Sub Text1_GotFocus(Index As Integer)
   Call adll.Enfoquetexto(Text1(Index))
'   Ctr_Ayutransf.Visible = False
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
  Dim rb As New ADODB.Recordset
  On Error Resume Next
  If KeyAscii = 13 Then
     If Index = 1 Then
         Set rb = VGCNx.Execute("select * from te_operaciongeneral where operacioncodigo='" & Escadena(Text1(1)) & "' and operacioncontrolaclienteprov='" & IIf(adll.ComboDato(Combo1) = "I", "P", "P") & "'")
         If rb.RecordCount > 0 Then
            Text1(1) = Escadena(rb!operacioncodigo)
            Label2(0) = Escadena(rb!operaciondescripcion)
 '           tipooperacion = 1
            controlarendicion = False
            If Escadena(rb!operacionvalidacajabancos) = "B" Then
                cayuda(1).Enabled = True
                Ctr_AyudaCaja.xclave = "": Label2(1) = ""
                Ctr_AyudaCaja.Visible = False
                cayuda(1).Enabled = False
                Text2(9).Enabled = True
                Text2(5).Enabled = True
                rb.Close
                Set rb = Nothing
                Ctr_Ayuda2.SetFocus
'                Combo2.SetFocus
 '               tipooperacion = 0
                Exit Sub
            Else
                Ctr_AyudaCaja.Visible = True
                cayuda(1).Enabled = True
                
                Text2(6).Enabled = False
                cayuda(5).Enabled = False
                Text2(5).Enabled = False
                Text2(9).Enabled = False
                cayuda(7).Enabled = False
                
                Ctr_AyudaCaja.SetFocus
                
                Set rb = Nothing
                Exit Sub
            End If
         Else
            Ctr_AyudaCaja.Visible = True
            cayuda(1).Enabled = True
            Text1(1) = "": Label2(0) = "": Ctr_AyudaCaja.xclave = "": Label2(1) = ""
         End If
         rb.Close
         Set rb = Nothing
     ElseIf Index = 2 Then
        Set rb = VGCNx.Execute("select * from te_codigocaja where cajacodigo='" & Ctr_AyudaCaja.xclave & "'")
        If rb.RecordCount > 0 Then
            Ctr_AyudaCaja.xclave = Escadena(rb!cajacodigo)
            Label2(1) = Escadena(rb!cajadescripcion)
        Else
            Ctr_AyudaCaja.xclave = ""
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
            MsgBox "El Proveedor no existe...Verifique!!", vbInformation, MsgTitle
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
  If Index = 73 Then
     Text2(3).Text = UCase(Text2(3).Text)
  End If
  
  If Index = 3 Then
     Text2(3).Text = UCase(Text2(3).Text)
  End If

End Sub

Private Sub Text2_LostFocus(Index As Integer)
 Dim i As Integer
 Dim lEncontro As Boolean
 Dim rs As ADODB.Recordset
 Dim SQL As String
 Dim VSnro As String
 If Index = 3 And VGParametros.empresaretencion = 1 And adll.ComboDato(Combo1) = "E" Then
    Dim rb As New ADODB.Recordset
   Set rb = VGCNx.Execute("select * from cp_cargo where clientecodigo='" & Ctr_Ayuda2.xclave & "' and documentocargo='" & Text2(1) & "' and cargonumdoc='" & Text2(11).Text & Text2(2).Text & "'")
   If rb.RecordCount() > 0 Then
      Emiteretencion = IIf(IsNull(rb!cargoemiteretencion), 0, rb!cargoemiteretencion)
     If rb!cargoemitedetraccion = 1 Then MsgBox "Documento tiene Detraccion", vbInformation, Caption
  Else
     Emiteretencion = 0
  End If
  If Emiteretencion = 1 Then Call detalleretencion
End If
If Index = 8 Then
      SQL = "select monedacodigo,isnull(cargoapeimpape,0)-isnull(cargoapeimppag,0) from cp_cargo "
      SQL = SQL & "where empresacodigo='" & Ctr_Ayuempresa.xclave & "' and documentocargo='" & Text2(1).Text & "' "
      SQL = SQL & " and cargonumdoc='" & Trim(Text2(11).Text & Text2(2).Text) & "' and clientecodigo='" & Trim(Ctr_Ayuda2.xclave) & "'"
      SQL = SQL & "and isnull(cargoapeflgreg,0) <> 1 "
      Set rs = VGCNx.Execute(SQL)
      If Not rs.BOF And Not rs.EOF Then
        If Text2(7).Text = rs(0) Then
          If Round(numero(Text2(8).Text), 2) > Round(rs(1), 2) Then
            MsgBox "El Monto a Pagar es mayor que el Saldo del Documento, valor=  -- > " & Round(rs(1), 2) & "", vbInformation, Caption
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
   
   lEncontro = False
   If rsdetat.RecordCount > 0 Then
      If Index = 6 And Text2(5).Text <> Empty And Len(Trim(Ctr_AyudaCaja.xclave)) = 0 Then
         rsdetat.MoveFirst
         Do Until rsdetat.EOF
           If Trim(rsdetat.Fields(6).Value) <> Trim(Text2(6).Text) And rsdetat.Fields(4).Value <> VGParametros.empresacodigoretencion Then
             lEncontro = True
             Exit Do
           End If
           rsdetat.MoveNext
         Loop
         If lEncontro = True Then
            MsgBox "El N� de Cheque debe ser el mismo para los Documentos", vbInformation, Caption
            Text2(6).Text = Empty
            Text2(6).SetFocus
         End If
      End If
   End If
   lEncontro = False
   '***JCGI
   If Index = 6 And Text2(5).Text <> Empty And Len(Trim(Ctr_AyudaCaja.xclave)) = 0 Then
    If Trim(Text2(9).Text) = Empty Then
        MsgBox "Debe ingresar N� de cuenta corriente", vbInformation, Caption
        Text2(6).Text = Empty
        Text2(9).SetFocus
    Else
      Set rs = New ADODB.Recordset
      SQL = "select count(*) from te_detallerecibos "
      SQL = SQL & "where detrec_emisioncheque='B' and detrec_cajabanco1='" & Text2(5).Text & "' AND "
      SQL = SQL & "detrec_monedacancela='" & Text2(7).Text & "' AND detrec_ndqc='" & Text2(6).Text & "' AND "
      SQL = SQL & "detrec_tdqc='" & Text2(4).Text & "' AND isnull(detrec_estadoreg,0)<>'1' AND detrec_numctacte='" & Text2(9).Text & "'"
      Set rs = VGCNx.Execute(SQL)
      If rs(0) > 0 Then
         If MsgBox("El N� de Cheque ya fue registrado , desea continuar (S/N)") = vbNo Then
            Text2(6).Text = Empty
            Text2(6).SetFocus
         End If
      End If
      Set rs = Nothing
    End If
   End If
   lEncontro = False
If Index = 2 Then
     If Len(Text2(Index)) = 0 Then Exit Sub
     Call Text2_KeyPress(Index, 13)
End If
End Sub


Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
 Dim rb As New ADODB.Recordset
 Dim rb1 As New ADODB.Recordset
 Dim VSnro As String
  If KeyAscii = 13 Then
    Text2(Index) = UCase(Text2(Index))
    If Index = 1 Then
        Set rb = VGCNx.Execute("select * from cp_tipodocumento where tdocumentoingplan='1' and tdocumentocodigo='" & Text2(1) & "'")
        If rb.RecordCount = 0 Then
          MsgBox "No existe tipo de documento...Verifique!!", vbInformation, MsgTitle
          rb.Close
          Set rb = Nothing
          Exit Sub
         ElseIf rb!tdocumentonumeauto = 1 Then
              Text2(11).Text = Left(rb!tdocumentonumerador, 4)
              Text2(2).Text = Mid(rb!tdocumentonumerador, 5, Len(rb!tdocumentonumerador) - 3)
        End If
        rb.Close
        Set rb = Nothing
    ElseIf Index = 2 Then
           If IsNumeric(Text2(Index)) Then
              VSnro = Right("0000000000" & Trim(Text2(Index)), Text2(Index).MaxLength)
           End If
        SQL = "select * from cp_cargo where empresacodigo='" & Ctr_Ayuempresa.xclave & "'"
        SQL = SQL & " and clientecodigo='" & Ctr_Ayuda2.xclave & "' and documentocargo='" & Text2(1) & "' and cargonumdoc='" & Text2(11).Text & VSnro & "'"
        SQL = SQL & "and isnull(cargoapeflgreg,0)<>1 "
        Set rb = VGCNx.Execute(SQL)
        If adll.ComboDato(Combo1.Text) = "I" Then
          If rb.RecordCount > 0 Then
             MsgBox "Existe el N� Documento Referenciado...Verifique!!", vbInformation, MsgTitle
             Text2(2).SetFocus
             Exit Sub
          End If
       Else
          If rb.RecordCount = 0 Then
             Set rb1 = VGCNx.Execute("select * from cp_tipodocumento where tdocumentoingplan='1' and tdocumentocodigo='" & Text2(1) & "' and tdocumentotipo='A'")
             If rb1.RecordCount = 0 Then
                MsgBox "No Existe el tipo de documento ...Verifique!!", vbInformation, MsgTitle
                Text2(1).SetFocus
                Exit Sub
             End If
          End If
       End If
       rb.Close
       Text2(Index) = VSnro
    ElseIf Index = 3 Then
      Text2(Index) = UCase(Text2(Index))
      If Not Text2(3) Like "[TP]" Then
        MsgBox "Solo debe ingresar P � T", vbInformation, MsgTitle
        Exit Sub
      End If
    ElseIf Index = 4 Or Index = 9 Then 'Tipo de cancelacion
      Set rb = VGCNx.Execute("select * from cp_tipodocumento where tdocumentotipo='A' and tdocumentoingcobra='1' and tdocumentocodigo='" & Text2(4) & "'")
      If rb.RecordCount = 0 Then
        MsgBox "No existe tipo de documento...Verifique!!", vbInformation, MsgTitle
        rb.Close
        Set rb = Nothing
        Exit Sub
       ElseIf rb!tdocumentonumeauto = 1 Then
           Text2(6).Text = rb!tdocumentonumerador
          Else
          Text2(6).Text = ""
          If rb!tdocumentovalidabanco = 1 Then
             Text2(5).Enabled = True
             Text2(6).Enabled = True
             cayuda(5).Enabled = True
           Else
             Text2(5).Enabled = False
             Text2(6).Enabled = False
             cayuda(5).Enabled = False

          End If
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
        MsgBox "El importe debe ser mayor que cero. Se corregir� el importe", vbInformation, "Aviso"
        Text2(8) = numero(Text2(8) * (-1))
       End If
  
    ElseIf Index = 9 Then
      Set rb = VGCNx.Execute("select * from te_cuentabancos inner join gr_banco on te_cuentabancos.cbanco_codigo=gr_banco.bancocodigo where gr_banco.bancocodigo='" & Text2(5) & "' and te_cuentabancos.monedacodigo='" & Text2(7) & "' and te_cuentabancos.cbanco_numero='" & Trim(Text2(9)) & "'")
      If rb.RecordCount = 0 Then
        MsgBox "No existe la cuenta corriente del banco indicado .... Verifique!!", vbInformation, MsgTitle
        rb.Close
        Set rb = Nothing
        Exit Sub
      End If
      rb.Close
      Set rb = Nothing
    ElseIf Index = 10 Then
           If Len(Trim(Ctr_AyudaCaja.xclave)) = 0 Then
              Call grabacion
              Exit Sub
            End If
    ElseIf Index = 11 Then
           If IsNumeric(Text2(Index)) Then
              Text2(Index) = Right("0000" & Trim(Text2(Index)), Text2(Index).MaxLength)
           End If
    ElseIf Index = 12 Then
           Text2(6).Text = Text2(12).Text
    End If
    Call Seguir(Text2(Index), KeyAscii)
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

Function aplicaciones()
  Dim acmd As New ADODB.Command
  Dim rb As New ADODB.Recordset
  Dim xabono, xzona, xmone, xcuenta, xtipo As String
  Dim xnumplan, ximpsol, xtcam, xnumpag As Double

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
                         
                         Set rb = VGCNx.Execute("select * from cp_cargo where documentocargo='" & rsdetat.Fields(4) & "' and cargonumdoc='" & rsdetat.Fields(6) & "'")
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
                                         
                         Set acmd.ActiveConnection = VGGeneral
                         acmd.CommandType = adCmdStoredProc
                         acmd.CommandText = "cp_abonadocumento_pro"
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
                             .Parameters("@ctabanco") = Trim(rsdetat.Fields(10))      'Cuenta Banco
                             .Parameters("@abonoflpres") = "1"
                             .Parameters("@abonocanimpcan") = CDbl(rsdetat.Fields(8))
                             .Parameters("@abonocanimpsol") = ximpsol
                             .Parameters("@usuario") = VGUsuario
                             .Parameters("@fechaact") = Date
                             .Parameters("@forma") = rsdetat.Fields(3)
                             .Parameters("@monedacan") = rsdetat.Fields(7)
                             .Parameters("@abonocantd") = rsdetat.Fields(1)
                             .Parameters("@abonocannro") = Trim(rsdetat.Fields(2))
                             .Parameters("@fechacan") = rsdetat.Fields(9)
                             .Parameters("@cliente") = Escadena(Ctr_Ayuda2.xclave)
                         End With
                         acmd.Execute
                         
                         Set acmd = Nothing
                         DoEvents
                                         
                         '**** Actualizamos Saldos de documento pendiente
                         If rsdetat.Fields(7) = g_tipodolar Then
                            If xmone = g_tiposol Then
                                    VGCNx.Execute "Update  cp_cargo Set cargoapeimppag=isnull(cargoapeimppag,0)+" & CDbl(rsdetat.Fields(8) * xtcam) & "," & _
                                             " cargoapenumpag='" & xnumpag + 1 & "'" & _
                                            " Where documentocargo='" & rsdetat.Fields(4) & "' and cargonumdoc='" & Trim(rsdetat.Fields(6)) & "' and " & _
                                            " clientecodigo='" & Trim(Ctr_Ayuda2.xclave) & "'"
                                            
                            Else
                                     VGCNx.Execute "Update cp_cargo Set cargoapeimppag=isnull(cargoapeimppag,0)+" & CDbl(rsdetat.Fields(8)) & "," & _
                                                " cargoapenumpag='" & xnumpag + 1 & "'" & _
                                                " Where documentocargo='" & rsdetat.Fields(4) & "' and cargonumdoc='" & Trim(rsdetat.Fields(6)) & "' and " & _
                                                " clientecodigo='" & Trim(Ctr_Ayuda2.xclave) & "'"
                                
                            End If
                         ElseIf rsdetat.Fields(7) = g_tiposol Then
                            If xmone = g_tipodolar Then
                                VGCNx.Execute "Update  cp_cargo Set cargoapeimppag=isnull(cargoapeimppag,0)+" & CDbl(rsdetat.Fields(8) / xtcam) & "," & _
                                           " cargoapenumpag='" & xnumpag + 1 & "'" & _
                                           " Where documentocargo='" & rsdetat.Fields(4) & "' and cargonumdoc='" & Trim(rsdetat.Fields(6)) & "' and " & _
                                           " clientecodigo='" & Trim(Ctr_Ayuda2.xclave) & "'"
                            Else
                                VGCNx.Execute "Update  cp_cargo Set cargoapeimppag=isnull(cargoapeimppag,0)+" & CDbl(rsdetat.Fields(8)) & "," & _
                                           " cargoapenumpag='" & xnumpag + 1 & "'" & _
                                           " Where documentocargo='" & rsdetat.Fields(4) & "' and cargonumdoc='" & Trim(rsdetat.Fields(6)) & "' and " & _
                                           " clientecodigo='" & Trim(Ctr_Ayuda2.xclave) & "'"
                            End If
                         End If
                         
                         VGCNx.Execute "Update  cp_cargo " & _
                                     " Set cargoapeflgcan= CASE Round(isnull(cargoapeimpape,0),2)-Round(isnull(cargoapeimppag,0),2) WHEN 0 THEN '1' ELSE '0' END ," & _
                                     "   cargoapefeccan='" & rsdetat.Fields(9) & "'" & _
                                     " Where documentocargo='" & rsdetat.Fields(4) & "' and cargonumdoc='" & Trim(rsdetat.Fields(6)) & "' and " & _
                                     " clientecodigo='" & Trim(Ctr_Ayuda2.xclave) & "'"
                         
                         '**** Actualizamos Saldos del cliente
                         If rsdetat.Fields(7) = g_tipodolar Then
                               VGCNx.Execute "Update  cp_proveedor Set clientesaldodolares=isnull(clientesaldodolares,0)-" & CDbl(rsdetat.Fields(8)) & _
                                           " Where clientecodigo='" & Ctr_Ayuda2.xclave & "'"
                         ElseIf rsdetat.Fields(7) = g_tiposol Then
                               VGCNx.Execute "Update  cp_proveedor Set clientesaldosoles=isnull(clientesaldosoles,0)-" & CDbl(rsdetat.Fields(8)) & _
                                           " Where clientecodigo='" & Ctr_Ayuda2.xclave & "'"
                         End If
             
                  End If
             End If

End Function

Public Function detalleretencion()
FrameRetencion.Visible = True
TxtMonto.Text = Text2(8).Text
Text2(18).Text = VGParametros.empresacodigoretencion
TxtMonto.SetFocus
End Function

Private Sub TxtMonto_KeyPress(KeyAscii As Integer)
Dim rb As New ADODB.Recordset
Set rb = VGCNx.Execute("select * from cp_cargo where clientecodigo='" & Ctr_Ayuda2.xclave & "' and documentocargo='" & Text2(1) & "' and cargonumdoc='" & Text2(11).Text & Text2(2).Text & "'")
If rb.RecordCount() > 0 Then
   Emiteretencion = IIf(IsNull(rb!cargoemiteretencion), 0, rb!cargoemiteretencion)
 Else
   Emiteretencion = 0
End If
If KeyAscii = 13 Then
   If Emiteretencion = 1 Then
      TxtMontopagar.Text = Round(TxtMonto.Text * (1 - VGParametros.porcentajeretencion / 100), 2)
      TxtMontoretencion.Text = Round(TxtMonto.Text * (VGParametros.porcentajeretencion) / 100, 2)
    Else
      TxtMontopagar.Text = TxtMonto.Text
      TxtMontoretencion.Text = 0
  End If
End If
rb.Close
Set rb = Nothing
'Text2(9).TextSetFocus
End Sub
Private Sub adicionaretencion()
Dim rsdetx As New ADODB.Recordset
Dim i As Integer
Dim n As Integer
Dim nroretencion As String
Dim ximporte As Double
Dim xmonto As Double
On Error GoTo xerror
 rsdetat.UpdateBatch adAffectAllChapters
Set rsdetx = VGCNx.Execute(" select * from " & Tabla)
Set rsdetat = VGCNx.Execute(" delete " & Tabla)
SQL = "select * from " & Tabla
rsdetat.Open SQL, VGCNx, adOpenDynamic, adLockBatchOptimistic
TDBGrid1.DataSource = rsdetat
rsdetx.MoveFirst
n = 0
Do Until rsdetx.EOF
   n = n + 1
   rsdetat.AddNew
   For i = 0 To rsdetx.Fields.Count - 1
    rsdetat.Collect(rsdetx.Fields(i).Name) = rsdetx.Collect(i)
   Next
   xmonto = rsdetx!importe
   If rsdetx!retencion = 0 Then
      rsdetat!importe = Round((1 - VGParametros.porcentajeretencion / 100) * rsdetx!importe, 2)
   End If
   rsdetat!Item = n
   If rsdetx!tdqc <> VGParametros.empresacodigoretencion And rsdetx!retencion = 0 Then
      n = n + 1
      rsdetat.AddNew
      For i = 0 To rsdetx.Fields.Count - 1
         rsdetat.Collect(rsdetx.Fields(i).Name) = rsdetx.Collect(i)
      Next
      rsdetat!importe = rsdetx!importe - Round((1 - VGParametros.porcentajeretencion / 100) * rsdetx!importe, 2)
      rsdetat!Item = n
      rsdetat!tdqc = VGParametros.empresacodigoretencion
      rsdetat!detrec_ndqc = nroretencion
      rsdetat!retencion = "1"
     Else
       If rsdetx!tdqc = VGParametros.empresacodigoretencion Then nroretencion = rsdetx!detrec_ndqc
     End If
   rsdetx.MoveNext
Loop
TDBGrid1.Refresh
Exit Sub
xerror:
  
  MsgBox "No se pudo Grabar " & Err.Description & " - " & Err.Number, vbInformation, Caption
End Sub
Private Sub Actualizasaldorendicion()
Dim rsaux As New ADODB.Recordset
Dim xrendicion As String
Dim xsaldo As Double
SQL = "select numero=max(rendicionnumero) from te_rendiciones where oficinacodigo='" & VGoficina & "'"
SQL = SQL & " and codigocaja='" & Ctr_AyudaCaja.xclave & "'"
Set rsaux = VGCNx.Execute(SQL)
xrendicion = rsaux!numero
If Ctr_AyudaCaja.xclave = "02" Then
   xsaldo = Label5(0) * Text1(3).Text
 Else
   xsaldo = Label5(0)
End If
SQL = " update te_rendiciones set saldoacumuladoxrendir=isnull(saldoacumuladoxrendir,0)-" & xsaldo
SQL = SQL & " where oficinacodigo='" & VGoficina & "'"
SQL = SQL & " and codigocaja='" & Ctr_AyudaCaja.xclave & "' and rendicionnumero='" & xrendicion & "'"
Set rsaux = VGCNx.Execute(SQL)
End Sub

Private Function VerificaNumDocIngresado() As Boolean
Dim filas, r As Integer
    VerificaNumDocIngresado = False
    filas = rsdetat.RecordCount
    If filas > 0 Then
        For r = 0 To filas - 1
            rsdetat.MoveFirst
            If rsdetat.Fields(2) = Format(Escadena(Text2(11).Text), "0000") & Format(Escadena(Text2(2)), "0000000000") Then VerificaNumDocIngresado = True
            rsdetat.MoveNext
        Next r
    End If
End Function
