VERSION 5.00
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form FrmNotas 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de Notas de Ventas"
   ClientHeight    =   7095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9840
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   9840
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   2535
      Left            =   30
      TabIndex        =   30
      Top             =   -30
      Width           =   9735
      Begin VB.CheckBox chkInafecto 
         Alignment       =   1  'Right Justify
         Caption         =   "Inaf."
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
         Left            =   2490
         TabIndex        =   50
         Top             =   1740
         Width           =   735
      End
      Begin VB.CommandButton cAyuda 
         Caption         =   "..."
         Height          =   285
         Left            =   3450
         TabIndex        =   31
         Top             =   2100
         Width           =   255
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   8130
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1320
         Width           =   1425
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1110
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1320
         Width           =   1665
      End
      Begin MSMask.MaskEdBox MBox1 
         Height          =   285
         Index           =   0
         Left            =   1740
         TabIndex        =   32
         Top             =   -330
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   12648447
         Enabled         =   0   'False
         MaxLength       =   6
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MBox1 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "d/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   3
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   6240
         TabIndex        =   33
         Top             =   210
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   12648447
         AutoTab         =   -1  'True
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MBox1 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   8580
         TabIndex        =   34
         Top             =   210
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   12648447
         Enabled         =   0   'False
         MaxLength       =   8
         PromptChar      =   "_"
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuda1 
         Height          =   345
         Left            =   1110
         TabIndex        =   1
         Top             =   930
         Width           =   8475
         _ExtentX        =   14949
         _ExtentY        =   609
         XcodMaxLongitud =   11
         xcodwith        =   800
         NomTabla        =   "vt_Cliente"
         TituloAyuda     =   "Ayuda de Clientes"
         ListaCampos     =   $"FrmNotas.frx":0000
         XcodCampo       =   "clientecodigo"
         XListCampo      =   "clienterazonsocial"
         ListaCamposDescrip=   "Codigo,Descripcion,Ruc,Direccion,Distrito,LimiteCred,Saldo,T,P,M,D"
         ListaCamposText =   $"FrmNotas.frx":00E6
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuda3 
         Height          =   315
         Left            =   5370
         TabIndex        =   13
         Top             =   2070
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   556
         XcodMaxLongitud =   2
         xcodwith        =   100
         NomTabla        =   "cc_conceptos"
         TituloAyuda     =   "Ayuda de Conceptos"
         ListaCampos     =   "conceptocodigo(1),conceptodescripcion(1)"
         XcodCampo       =   "conceptocodigo"
         XListCampo      =   "conceptodescripcion"
         ListaCamposDescrip=   "Codigo,Descripcion"
         ListaCamposText =   "conceptocodigo,conceptodescripcion"
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
         Height          =   300
         Index           =   3
         Left            =   5880
         TabIndex        =   5
         Top             =   1335
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   529
         _Version        =   393216
         ClipMode        =   1
         AllowPrompt     =   -1  'True
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MBox 
         Height          =   315
         Index           =   1
         Left            =   2940
         TabIndex        =   3
         Top             =   1320
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   4
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MBox 
         Height          =   315
         Index           =   2
         Left            =   3480
         TabIndex        =   4
         Top             =   1320
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MBox 
         Height          =   255
         Index           =   6
         Left            =   1110
         TabIndex        =   10
         Top             =   2100
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   3
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MBox 
         Height          =   255
         Index           =   7
         Left            =   1560
         TabIndex        =   11
         Top             =   2100
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   3
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MBox 
         Height          =   255
         Index           =   8
         Left            =   2100
         TabIndex        =   12
         Top             =   2100
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   8
         PromptChar      =   "_"
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuda4 
         Height          =   315
         Left            =   1140
         TabIndex        =   0
         Top             =   210
         Width           =   3525
         _ExtentX        =   6218
         _ExtentY        =   556
         XcodMaxLongitud =   3
         xcodwith        =   200
         NomTabla        =   "co_multiempresas"
         TituloAyuda     =   "Ayuda de Empresas"
         ListaCampos     =   "empresacodigo(1),empresadescripcion(1)"
         XcodCampo       =   "empresacodigo"
         XListCampo      =   "empresadescripcion"
         ListaCamposDescrip=   "Codigo,Descripcion"
         ListaCamposText =   "empresacodigo,empresadescripcion"
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuda2 
         Height          =   315
         Left            =   6990
         TabIndex        =   9
         Top             =   1680
         Width           =   2595
         _ExtentX        =   4577
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
         Height          =   300
         Index           =   5
         Left            =   4770
         TabIndex        =   8
         Top             =   1695
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   529
         _Version        =   393216
         ClipMode        =   1
         AllowPrompt     =   -1  'True
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MBox 
         Height          =   315
         Index           =   4
         Left            =   1110
         TabIndex        =   7
         Top             =   1680
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         PromptChar      =   "_"
      End
      Begin VB.Label Label5 
         Caption         =   "Importe"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   7
         Left            =   120
         TabIndex        =   49
         Top             =   1740
         Width           =   825
      End
      Begin VB.Label Label5 
         Caption         =   "Fe. Vencimiento"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   3
         Left            =   3360
         TabIndex        =   48
         Top             =   1740
         Width           =   1395
      End
      Begin VB.Label Label5 
         Caption         =   "Vendedor"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   6090
         TabIndex        =   47
         Top             =   1710
         Width           =   1605
      End
      Begin VB.Label LblFecDoc 
         Height          =   285
         Left            =   3780
         TabIndex        =   46
         Top             =   2100
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.Label Label5 
         Caption         =   "Concepto"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   5
         Left            =   4290
         TabIndex        =   45
         Top             =   2130
         Width           =   1305
      End
      Begin VB.Label Label5 
         Caption         =   "Documento"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   4
         Left            =   90
         TabIndex        =   44
         Top             =   2100
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Moneda"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   7290
         TabIndex        =   43
         Top             =   1380
         Width           =   915
      End
      Begin VB.Label Label5 
         Caption         =   "Fe. Emision"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   4860
         TabIndex        =   42
         Top             =   1380
         Width           =   1035
      End
      Begin VB.Label Label4 
         Caption         =   "Cliente"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   41
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "DETALLE DOCUMENTO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   210
         TabIndex        =   40
         Top             =   630
         Width           =   3795
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         Index           =   1
         X1              =   30
         X2              =   9720
         Y1              =   570
         Y2              =   570
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000009&
         Index           =   0
         X1              =   30
         X2              =   9750
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label Label1 
         Caption         =   "Nro. Planilla"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   480
         TabIndex        =   39
         Top             =   -300
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo Cambio"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   7470
         TabIndex        =   38
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Documento"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   6
         Left            =   90
         TabIndex        =   37
         Top             =   1380
         Width           =   1395
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Registro"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   4890
         TabIndex        =   36
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Empresa :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   3
         Left            =   210
         TabIndex        =   35
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2985
      Left            =   45
      TabIndex        =   18
      Top             =   2550
      Width           =   9735
      Begin VB.Frame Frame3 
         Height          =   675
         Left            =   120
         TabIndex        =   20
         Top             =   2190
         Width           =   9465
         Begin VB.TextBox Text1 
            BackColor       =   &H00C0FFFF&
            Enabled         =   0   'False
            ForeColor       =   &H00404040&
            Height          =   285
            Index           =   3
            Left            =   8040
            MaxLength       =   12
            TabIndex        =   24
            Top             =   210
            Width           =   1245
         End
         Begin VB.TextBox Text1 
            BackColor       =   &H00C0FFFF&
            Enabled         =   0   'False
            ForeColor       =   &H00404040&
            Height          =   285
            Index           =   2
            Left            =   5910
            MaxLength       =   12
            TabIndex        =   23
            Top             =   240
            Width           =   1245
         End
         Begin VB.TextBox Text1 
            BackColor       =   &H00C0FFFF&
            Enabled         =   0   'False
            ForeColor       =   &H00404040&
            Height          =   285
            Index           =   1
            Left            =   3480
            MaxLength       =   12
            TabIndex        =   22
            Top             =   210
            Width           =   1035
         End
         Begin VB.TextBox Text1 
            BackColor       =   &H00C0FFFF&
            Enabled         =   0   'False
            ForeColor       =   &H00404040&
            Height          =   285
            Index           =   0
            Left            =   1410
            MaxLength       =   12
            TabIndex        =   21
            Top             =   240
            Width           =   1245
         End
         Begin VB.Label Label2 
            Caption         =   "TOTAL"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   3
            Left            =   7380
            TabIndex        =   28
            Top             =   270
            Width           =   675
         End
         Begin VB.Label Label2 
            Caption         =   "TOTAL IGV"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   2
            Left            =   4680
            TabIndex        =   27
            Top             =   270
            Width           =   1095
         End
         Begin VB.Label Label2 
            Caption         =   "IMPORTE"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   1
            Left            =   360
            TabIndex        =   26
            Top             =   270
            Width           =   1035
         End
         Begin VB.Label Label2 
            Caption         =   "IGV"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   2850
            TabIndex        =   25
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.TextBox Text2 
         Height          =   1635
         Left            =   150
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   19
         Top             =   510
         Width           =   9435
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "REFERENCIA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   180
         TabIndex        =   29
         Top             =   180
         Width           =   9405
      End
   End
   Begin VB.CommandButton cmdBotones 
      Caption         =   "&Nuevo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   990
      Index           =   0
      Left            =   3135
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5640
      Width           =   1110
   End
   Begin VB.CommandButton cmdBotones 
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   990
      Index           =   12
      Left            =   5775
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5640
      Width           =   1095
   End
   Begin VB.CommandButton cmdBotones 
      Caption         =   "&Acepta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   990
      Index           =   11
      Left            =   4455
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5640
      Width           =   1110
   End
   Begin MSComctlLib.StatusBar Panel 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   17
      Top             =   6750
      Width           =   9840
      _ExtentX        =   17357
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
   Begin Crystal.CrystalReport oCrystalReport 
      Left            =   180
      Top             =   4950
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "FrmNotas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adll As New dllgeneral.dll_general

Dim nLongicampo(6) As Integer
Dim rsdeta As New ADODB.Recordset
'FIXIT: Declare 'wCabe' con un tipo de datos de enlace en tiempo de compilación            FixIT90210ae-R1672-R1B8ZE
Dim wCabe(40)

Dim apedido As String
Dim aalmacen As String
Dim alista As String * 2

Private Sub chkInafecto_Click()
    Text1(0) = numero(MBox(4))
    If Len(RTrim$(Text1(1))) > 0 And chkInafecto.Value = 0 Then
        Text1(2) = RTrim$(numero(CDbl(Text1(0)) * CDbl(Text1(1)) / 100))
        Text1(3) = RTrim$(numero(CDbl(Text1(0)) + CDbl(Text1(2))))
    Else
        Text1(2) = "0"
        Text1(3) = RTrim$(numero(CDbl(numero(MBox(4)))))
    End If
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub Ctr_Ayuda4_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
VGparametros.empresacodigo = Ctr_Ayuda4.xclave
End Sub

Private Sub Form_Load()
MostrarForm Me, "C"
MBox1(1) = Format(Date, "DD/MM/YYYY")
  
Call Ctr_Ayuda1.Conexion(VGCNx)
Call Ctr_Ayuda2.Conexion(VGCNx)
Call Ctr_Ayuda3.Conexion(VGCNx)
Call Ctr_Ayuda4.Conexion(VGCNx)

Call adll.llenacombo(Combo2, "select * from cc_tipodocumento where tdocumentonotaconta=1 ", VGCNx)
Call adll.llenacombo(Combo1, "select * from gr_moneda", VGCNx)

MBox1(2) = Format(DatoTipoCambio(VGcnxCT, Date), "##0.00")
Text1(1) = (VGparametros.igv)
'cmdBotones(0).Picture = MDIPrincipal.ImgList2.ListImages.Item("Crear").Picture
'cmdBotones(11).Picture = MDIPrincipal.ImgList2.ListImages.Item("Grabar").Picture
'cmdBotones(12).Picture = MDIPrincipal.ImgList2.ListImages.Item("Retornar").Picture
  
End Sub

Public Function GrabarData() As Integer
Dim J As Integer
Dim regi As Long
Dim nsql As String
Dim ltipo As String
Dim lzona As String
Dim Previo As Double
Dim tinafecto As Double
Dim xserie As String * 3
Dim xfactu As String * 5
Dim xtipofac As String * 2
Dim fechasunat As Date
Dim tcargo As String
Dim RsSerie As New ADODB.Recordset
Dim acmd As New ADODB.Command
Dim asql As New ADODB.Recordset
Dim arbusca As New ADODB.Recordset
Dim existedoc As Integer
On Error GoTo vererror

GrabarData = 0
existedoc = 0
'******** CABECERA DE MOVIMIENTO *****************
For J = 1 To 29
    wCabe(J) = ""
Next J
fechasunat = MBox(3)
apedido = MBox(6) & RTrim$(MBox(7) & MBox(8))

Set asql = VGCNx.Execute("select * from vt_pedido where pedidotipofac='" & MBox(6) & "' and pedidonrofact='" & RTrim$(MBox(7) & MBox(8)) & "' and empresacodigo='" & VGparametros.empresacodigo & "' ")
If asql.RecordCount > 0 Then
   existedoc = 1
   apedido = Escadena(asql!pedidonumero)
   wCabe(1) = Escadena(asql!puntovtacodigo)         'Escadena(asql!p)                       'Pto Venta
   wCabe(2) = Escadena(asql!pedidonumero)           'rtrim$(MBox(1))                       'nro pedido
   wCabe(3) = Escadena(asql!pedidonrofact)          'rtrim$(MBox(2))                        'nro factura
   wCabe(4) = Escadena(asql!pedidonrofact)          'rtrim$(MBox(3))                         'nro boleta
   wCabe(5) = Escadena(asql!pedidonrofact)          'rtrim$(MBox(4))                         'nro guia
   wCabe(6) = 0      'MBox(5)                       'dscto gral
   wCabe(7) = 0      'MBox(6)                       'dscto promocional
   wCabe(8) = 0      'MBox(7)                       'dscto especial
   wCabe(9) = adll.ComboDato(Combo1.Text)           'moneda
   wCabe(10) = CDbl(MBox1(2))                       'tipo de cambio
   wCabe(11) = CDbl(Escadena(asql!pedidolistaprec)) 'dllgeneral.ComboDato(Combo2.Text)       'lista de precios
   wCabe(12) = " "                                  'MBox(9)                      'mensajes
   wCabe(13) = Escadena(asql!modovtacodigo)         'dllgeneral.ComboDato(Combo3.Text)       'modo de venta
   wCabe(14) = MBox1(1)                             'MBox(10)                     'fecha de atencion
   wCabe(15) = Escadena(asql!formapagocodigo)       'dllgeneral.ComboDato(Combo4.Text)       'forma de pago
   wCabe(16) = Ctr_Ayuda1.xclave                    'MBox(11)                     'cliente
   wCabe(17) = Ctr_Ayuda2.xclave                    'MBox(12)                       'vendedor
   wCabe(18) = 0    'MBox(13)                       'comision
   wCabe(19) = Escadena(asql!almacencodigo)         'Ctr_Ayuda3.xclave        'MBox(14)                     'almacen
   wCabe(20) = 0      'MBox(15)                     'otros gastos
   wCabe(21) = "0"      'MBox(16)                   'nota pedido
   wCabe(22) = "0"      'MBox(17)                   'orden de compra
   wCabe(23) = Escadena(asql!pedidoautorizacion)    'dllgeneral.ComboDato(Combo5.Text)       'autorizacion
   wCabe(24) = 0       'MBox(18)                    'dias pago
   wCabe(25) = 0                           'Total Cantidad
   wCabe(26) = Round(Text1(0), 2)                         'Total Bruto
   wCabe(27) = 0    'MBox2(8)              'total fletes --T.D.
   wCabe(28) = Round(Text1(2), 2)          'Total Igv
   wCabe(29) = Round(Text1(3), 2)         'Neto a Facturar
   wCabe(30) = Escadena(asql!pedidoentrega)    'MBox(19)                    'entrega pedido
   wCabe(31) = Escadena(asql!clienterazonsocial)  'MBox3(1)                    'nombre cliente
   wCabe(32) = Escadena(asql!clientedireccion)    'MBox3(3)                    'direccion
   wCabe(33) = Escadena(asql!ClienteRuc)  'MBox3(2)                    'ruc
   wCabe(34) = MBox(3)                  'Date                           'fechafactura
   wCabe(35) = 0                     'Total Descuentos Globales
   wCabe(36) = 0                     'Total Descuentos Cliente
   wCabe(37) = 0                     'Total Descuentos Oficina
   wCabe(38) = 0                     'Total Descuentos Item
   wCabe(39) = 0                     'Total Descuentos Linea
   wCabe(40) = 0                     'Total Descuentos x Promocion
   fechasunat = IIf(IsNull(asql!pedidofechasunat), MBox(3), asql!pedidofechasunat)
   asql.Close
   Set asql = Nothing
Else
     'apedido = "00000000000"
      wCabe(1) = g_ptoventa               'Escadena(asql!puntovtacodigo)         'Escadena(asql!p)                       'Pto Venta
      wCabe(6) = 0      'MBox(5)                       'dscto gral
      wCabe(7) = 0      'MBox(6)                       'dscto promocional
      wCabe(8) = 0      'MBox(7)                       'dscto especial
      wCabe(9) = adll.ComboDato(Combo1.Text)           'moneda
      wCabe(10) = CDbl(MBox1(2))                       'tipo de cambio
      wCabe(11) = 1    'CDbl(Escadena(asql!pedidolistaprec)) 'dllgeneral.ComboDato(Combo2.Text)       'lista de precios
      wCabe(12) = " "                                  'MBox(9)                      'mensajes
      wCabe(13) = "02"  'Escadena(asql!modovtacodigo)         'dllgeneral.ComboDato(Combo3.Text)       'modo de venta
      wCabe(14) = MBox1(1)                             'MBox(10)                     'fecha de atencion
      wCabe(15) = ""   'Escadena(asql!formapagocodigo)       'dllgeneral.ComboDato(Combo4.Text)       'forma de pago
      wCabe(16) = Ctr_Ayuda1.xclave                    'MBox(11)                     'cliente
      wCabe(17) = Ctr_Ayuda2.xclave                    'MBox(12)                       'vendedor
      wCabe(18) = 0    'MBox(13)                       'comision
      wCabe(19) = ""   'Escadena(asql!almacencodigo)         'Ctr_Ayuda3.xclave        'MBox(14)                     'almacen
      wCabe(20) = 0        'MBox(15)                     'otros gastos
      wCabe(21) = "0"      'MBox(16)                   'nota pedido
      wCabe(22) = "0"      'MBox(17)                   'orden de compra
      wCabe(23) = ""       'Escadena(asql!pedidoautorizacion)    'dllgeneral.ComboDato(Combo5.Text)       'autorizacion
      wCabe(24) = 0        'MBox(18)                    'dias pago
      wCabe(25) = 0                           'Total Cantidad
      wCabe(26) = Round(Text1(0), 2)                         'Total Bruto
      wCabe(27) = 0    'MBox2(8)              'total fletes --T.D.
      wCabe(28) = Round(Text1(2), 2)          'Total Igv
      wCabe(29) = Round(Text1(3), 2)         'Neto a Facturar
      wCabe(30) = ""   'Escadena(asql!pedidoentrega)    'MBox(19)                    'entrega pedido
      wCabe(31) = Ctr_Ayuda1.xnombre   'Escadena(asql!clienterazonsocial)  'MBox3(1)                    'nombre cliente
      wCabe(32) = ""   'Escadena(asql!clientedireccion)    'MBox3(3)                    'direccion
      wCabe(33) = ""   'Escadena(asql!clienteruc)  'MBox3(2)                    'ruc
      wCabe(34) = MBox(3)                  'Date                           'fechafactura
      wCabe(35) = 0                     'Total Descuentos Globales
      wCabe(36) = 0                     'Total Descuentos Cliente
      wCabe(37) = 0                     'Total Descuentos Oficina
      wCabe(38) = 0                     'Total Descuentos Item
      wCabe(39) = 0                     'Total Descuentos Linea
      wCabe(40) = 0                     'Total Descuentos x Promocion
      fechasunat = MBox(3)  'IIf(IsNull(asql!pedidofechasunat), MBox(3), asql!pedidofechasunat)
      Dim rs As New ADODB.Recordset
      Set rs = New ADODB.Recordset
      Set rs = VGCNx.Execute("select cargoapefecemi from vt_cargo where documentocargo='" & MBox(6) & "' and cargonumdoc='" & RTrim$(MBox(7) & MBox(8)) & "' and empresacodigo='" & VGparametros.empresacodigo & "'")
      If Not rs.BOF And Not rs.EOF Then
         fechasunat = rs(0)
      Else
         fechasunat = MBox(3)
      End If
  End If

  Set asql = Nothing
  

    
' ** Verificando Numeracion de Documentos *****
Set RsSerie = VGCNx.Execute("select puntovtadocserie from vt_puntovtadocumento where documentocodigo='" & g_tipoped & "' and puntovtacodigo='" & g_ptoventa & "' and empresacodigo='" & VGparametros.empresacodigo & "'")
If RsSerie.RecordCount > 0 Then g_pedserie = RsSerie!puntovtadocserie

Set RsSerie = VGCNx.Execute("select puntovtadocserie from vt_puntovtadocumento where documentocodigo='" & g_tipofac & "' and puntovtacodigo='" & g_ptoventa & "' and empresacodigo='" & VGparametros.empresacodigo & "'")
If RsSerie.RecordCount > 0 Then g_facserie = RsSerie!puntovtadocserie

'-----------------------------------------------
wCabe(2) = g_pedserie & Right$("000000000000" & TraeDataSerie("select puntovtadoccorr from vt_puntovtadocumento where documentocodigo='" & g_tipoped & "' and puntovtadocserie='" & g_pedserie & "' and puntovtacodigo='" & g_ptoventa & "' and empresacodigo='" & VGparametros.empresacodigo & "'", VGCNx), 8)
'wCabe(34) = Date                       'fechafactura
wCabe(3) = g_facserie & Right$("000000000000" & TraeDataSerie("select puntovtadoccorr from vt_puntovtadocumento where documentocodigo='" & adll.ComboDato(Combo2.Text) & "' and puntovtadocserie='" & g_facserie & "' and puntovtacodigo='" & g_ptoventa & "' and empresacodigo='" & VGparametros.empresacodigo & "'", VGCNx), 8)
wCabe(4) = adll.ComboDato(Combo2.Text)
wCabe(5) = "0"
If adll.VerificaDatoExistente(VGCNx, "select * from vt_pedido where pedidonrofact='" & MBox(1) & MBox(2) & "' and pedidotipofac='" & adll.ComboDato(Combo2.Text) & "' and empresacodigo='" & VGparametros.empresacodigo & "'") = 1 Then
   MsgBox "Ya existe el Documento: " & adll.ComboDato(Combo2.Text) & "-" & MBox(1) & MBox(2), vbInformation, MsgTitle
   GrabarData = 0
   Exit Function
End If

'*** Verifica Serie Documentos *****
Set asql = VGCNx.Execute("select puntovtadoccorr from vt_puntovtadocumento Where documentocodigo='" & g_tipoped & "' and puntovtacodigo='" & g_ptoventa & "' and puntovtadocserie='" & MBox(1) & "' and empresacodigo='" & VGparametros.empresacodigo & "'")
If asql.RecordCount > 0 Then
   wCabe(2) = g_pedserie & Right$("000000000000" & RTrim$(asql!puntovtadoccorr), 8)
End If
asql.Close
Set asql = Nothing
    
nsql = "Update vt_puntovtadocumento " & _
        " set puntovtadoccorr='" & Right$("00000000" & RTrim$(CStr(wCabe(2) + 1)), 8) & "'" & _
        " Where documentocodigo='" & g_tipoped & "' and puntovtacodigo='" & g_ptoventa & "' and puntovtadocserie='" & g_pedserie & "' and empresacodigo='" & VGparametros.empresacodigo & "'"
VGCNx.Execute nsql
    
nsql = "Update vt_puntovtadocumento " & _
       " set puntovtadoccorr='" & Right$("00000000" & RTrim$(CStr(CDbl(MBox(2).Text) + 1)), 8) & "'" & _
       " Where documentocodigo='" & adll.ComboDato(Combo2.Text) & "' and puntovtacodigo='" & g_ptoventa & "' and puntovtadocserie='" & MBox(1).Text & "' and empresacodigo='" & VGparametros.empresacodigo & "'"
VGCNx.Execute nsql
               
DoEvents

'**cambio de documentacion
wCabe(5) = 0
    
DoEvents
Set acmd.ActiveConnection = VGGeneral
acmd.CommandType = adCmdStoredProc
acmd.CommandText = "cc_ingresanota_pro"
acmd.CommandTimeout = 0
acmd.Prepared = True
With acmd
    .Parameters("@base") = VGCNx.DefaultDatabase
    .Parameters("@tabla") = "vt_pedido"
    .Parameters("@tipo") = "1"
    .Parameters("@puntovta") = wCabe(1)
    .Parameters("@numero") = wCabe(2)
    .Parameters("@factura") = RTrim$(MBox(1)) & RTrim$(MBox(2))    'wCabe(3)
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
    .Parameters("@montodsctoppago") = 0
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
    .Parameters("@tiporefe") = Ctr_Ayuda3.xclave
    .Parameters("@nrorefe") = RTrim$(MBox(7) & MBox(8))
    .Parameters("@fsunat") = fechasunat
    .Parameters("@empresa") = VGparametros.empresacodigo
End With
acmd.Execute
Set acmd = Nothing
DoEvents
    
'*Grabar en los cargos ***ctacte ***
 lzona = "00"
 Set asql = VGCNx.Execute("select * from vt_zonavendedor where vendedorcodigo='" & wCabe(17) & "'")
 If asql.RecordCount > 0 Then
     lzona = Escadena(asql!zonacodigo)
 End If
 asql.Close
 Set asql = Nothing
    
 ltipo = "1"
 If adll.VerificaDatoExistente(VGCNx, "select * from vt_cargo where documentocargo='" & adll.ComboDato(Combo2.Text) & "' and cargonumdoc='" & RTrim$(MBox(1) & MBox(2)) & "' and empresacodigo='" & VGparametros.empresacodigo & "'") = 0 Then
   ltipo = "1"
 Else
   ltipo = "2"
 End If
 
 If adll.VerificaDatoExistente(VGCNx, "select * from cc_tipodocumento where tdocumentocodigo='" & adll.ComboDato(Combo2.Text) & "' and tdocumentotipo='A' ") = 1 Then
   tcargo = "A"
 Else
   tcargo = "C"
 End If


 Set acmd.ActiveConnection = VGGeneral
 acmd.CommandType = adCmdStoredProc
 acmd.CommandTimeout = 0
 acmd.CommandText = "cc_ingresacargovalor_pro"
 acmd.Prepared = True
 With acmd
     .Parameters("@base") = VGCNx.DefaultDatabase
     .Parameters("@tipo") = ltipo
     .Parameters("@tabla") = "vt_cargo"
     .Parameters("@tipodocu") = adll.ComboDato(Combo2.Text)
     .Parameters("@numero") = RTrim$(MBox(1) & MBox(2))
     .Parameters("@cliente") = Escadena(RTrim$(wCabe(16)))
     .Parameters("@vendedor") = Escadena(wCabe(17))
     .Parameters("@zona") = lzona
     .Parameters("@apefecemi") = wCabe(14)
     .Parameters("@moneda") = Escadena(wCabe(9))
     .Parameters("@apeimppag") = wCabe(29)
     .Parameters("@usuario") = g_usuario
     .Parameters("@tipocambio") = wCabe(10)
     .Parameters("@fechaact") = MBox1(1)
     .Parameters("@flagcancel") = "0"
     .Parameters("@cargoabono") = tcargo
     .Parameters("@referencia") = Left$(RTrim$(Text2), 254)
     .Parameters("@concepto") = Escadena(Ctr_Ayuda3.xclave)
     .Parameters("@venci") = MBox(3)
     .Parameters("@empresa") = VGparametros.empresacodigo
 End With
 acmd.Execute
 Set acmd = Nothing
     
MsgBox "Se Grabo Satisfactoriamente el Documento " & Chr(13) & Chr(10) & adll.ComboDato(Combo2.Text) & " >= " & MBox(1) & MBox(2), vbInformation, MsgTitle
GrabarData = 1
Exit Function
    
vererror:
   If Err Then
      MsgBox Err.Number & "-" & Err.Description
      MsgBox "Comunicarse con Sistemas ...!!" & Chr(13) & Chr(10) & VGCNx.Errors(0).Number & "-" & VGCNx.Errors(0).Description
      Exit Function
      Resume
   End If
End Function

Private Sub cAyuda_Click()
 nAyuda = "": nDetalle = ""
 If Len(RTrim$(MBox(6))) > 0 And Len(RTrim$(MBox(7))) > 0 And Len(RTrim$(MBox(8))) > 0 Then
    SendKeys "{tab}"
    Exit Sub
 End If
 
 If adll.VerificaDatoExistente(VGCNx, "select * from vt_pedido where clientecodigo='" & RTrim$(Ctr_Ayuda1.xclave) & "' and empresacodigo='" & Ctr_Ayuda4.xclave & "'") = 1 Then
       Dim gfiltra(2, 2) As String
       gfiltra(1, 1) = g_tipofac: gfiltra(1, 2) = "pedidonrofact"
       gfiltra(2, 1) = g_tipobol: gfiltra(2, 2) = "pedidonroboleta"
       FrmAyudaCli.TipoForma = 1
       FrmAyudaCli.BConexion = VGCNx   'cn
       FrmAyudaCli.Bdata = "0"
       FrmAyudaCli.BTabla = "vt_pedido"
       FrmAyudaCli.BCampos = "pedidotipofac as Tipo,pedidonrofact as Documento,pedidofecha as Fecha,pedidomoneda as Moneda,pedidototneto as Total"
       FrmAyudaCli.BOrden = "pedidofecha"
       FrmAyudaCli.BCondi = "clientecodigo='" & Ctr_Ayuda1.xclave & "' and empresacodigo='" & Ctr_Ayuda4.xclave & "'" '  and pedidotipofac='" & MBox(6).Text & "' "
       FrmAyudaCli.BFiltro = gfiltra
 Else
  If adll.VerificaDatoExistente(VGCNx, "select * from vt_cargo where clientecodigo='" & RTrim$(Ctr_Ayuda1.xclave) & "'  and cargoapecarabo='C' and isnull(cargoapeflgreg,0)<>1 and empresacodigo='" & VGparametros.empresacodigo & "'") = 1 Then
       Dim ffiltra(1, 1) As String
       ffiltra(1, 1) = g_tipofac: ffiltra(1, 2) = "cargonumdoc"
       FrmAyudaCli.TipoForma = 1
       FrmAyudaCli.BConexion = VGCNx   'cn
       FrmAyudaCli.Bdata = "0"
       FrmAyudaCli.BTabla = "vt_cargo"
       FrmAyudaCli.BCampos = "documentocargo as Tipo,cargonumdoc as Documento,cargoapefecemi as Fecha,monedacodigo as Moneda,cargoapeimpape as Total"
       FrmAyudaCli.BOrden = "cargoapefecemi"
       FrmAyudaCli.BCondi = "clientecodigo='" & RTrim$(Ctr_Ayuda1.xclave) & "' and empresacodigo='" & Ctr_Ayuda4.xclave & "' and cargoapecarabo='C' and isnull(cargoapeflgreg,0)<>1"
       FrmAyudaCli.BFiltro = ffiltra
   Else
       nAyuda = "": nDetalle = ""
       MsgBox "No existen documentos pendientes...", vbInformation, MsgTitle
       Exit Sub
   End If
 End If
 FrmAyudaCli.Show 1
 If Len(Escadena(nAyuda)) > 0 Then
    MBox(6) = Escadena(nAyuda): MBox(7) = Left$(Escadena(nDetalle), 3): MBox(8) = Right$(Escadena(nDetalle), 8)
    LblFecDoc.Caption = nfecha
 End If
 nAyuda = "": nDetalle = ""

End Sub

Private Sub cmdBotones_Click(Index As Integer)
   Dim asql As String
   Dim acmd As New ADODB.Command
'FIXIT: Declare 'J' con un tipo de datos de enlace en tiempo de compilación                FixIT90210ae-R1672-R1B8ZE
   Dim J, nl As Integer
   Dim nflag As Integer
   
   On Error GoTo vererror
   
   nflag = 0
   Select Case Index
    Case 0
       MBox(1) = "": MBox(2) = "": MBox(4) = "": MBox(6) = "": MBox(7) = "": MBox(8) = ""
       Text2 = ""
       'Limpiartexto Text1, 0, 3
       Text1(0).Text = Empty
       Text1(2).Text = Empty
       Text1(3).Text = Empty
       
       Ctr_Ayuda1.SetFocus
    
    Case 11
        If Len(RTrim$(Combo2.Text)) = 0 Then
            MsgBox "Falta seleccionar documento.", vbInformation, "Sistemas"
            Combo2.SetFocus
            Exit Sub
        End If
        If Len(RTrim$(MBox(1).Text)) = 0 Or Len(RTrim$(MBox(2).Text)) = 0 Then
            MsgBox "Nro de serie de nota no valido.", vbInformation, "Sistemas"
            Combo2.SetFocus
            Exit Sub
        End If
        If Not IsDate(MBox(3)) Then
            MsgBox "La fecha de emision no es correcta.", vbInformation, "Sistemas"
            MBox(3).SetFocus
            Exit Sub
        End If
        If Len(RTrim$(Combo1.Text)) = 0 Then
            MsgBox "Falta seleccionar moneda.", vbInformation, "Sistemas"
            Combo1.SetFocus
            Exit Sub
        End If
        
        If Len(RTrim$(MBox(4))) <> 0 Then
            If Not IsNumeric(MBox(4)) Then
                MsgBox "El importe ingresado no es numerico.", vbInformation, "Sistemas"
                MBox(4).SetFocus
                Exit Sub
            End If
        Else
            MsgBox "Falta ingresar el importe.", vbInformation, "Sistemas"
            MBox(4).SetFocus
            Exit Sub
        End If
        
        If Not IsDate(MBox(5)) Then
            MsgBox "La fecha de vencimiento no es correcta.", vbInformation, "Sistemas"
            MBox(5).SetFocus
            Exit Sub
        End If
        If IsNull(Ctr_Ayuda2.xclave) Or Len(RTrim$(Ctr_Ayuda2.xclave)) = 0 Then
           MsgBox "No existe Vendedor ...Verifique!!!", vbInformation, MsgTitle
           Ctr_Ayuda2.SetFocus
           Exit Sub
        End If
        If Len(RTrim$(MBox(6).Text)) = 0 Then
            MsgBox "Falta ingresar codigo de documento.", vbInformation, "Sistemas"
            MBox(6).SetFocus
            Exit Sub
        End If
        If IsNull(Ctr_Ayuda3.xclave) Or Len(RTrim$(Ctr_Ayuda3.xclave)) = 0 Then
           MsgBox "Codigo de conceptos no existe...Verifique!!!", vbInformation, MsgTitle
           Ctr_Ayuda3.SetFocus
           Exit Sub
        End If
        If Len(RTrim$(MBox(7).Text)) = 0 Or Len(RTrim$(MBox(8).Text)) = 0 Then
            MsgBox "El Nro de documento al cual se hace referencia no es valido.", vbInformation, "Sistemas"
            MBox(7).SetFocus
            Exit Sub
        End If
        If Ctr_Ayuda4.xclave = "" Then
            MsgBox "Falta seleccionar empresa", vbInformation, "Sistema"
            Ctr_Ayuda4.SetFocus
            Exit Sub
        End If

        If IsNull(Ctr_Ayuda1.xclave) Or Len(RTrim$(Ctr_Ayuda1.xclave)) = 0 Then
           MsgBox "Cliente no existe...Verifique!!!", vbInformation, MsgTitle
           Ctr_Ayuda1.SetFocus
           Exit Sub
        End If
        
        If IsNull(MBox1(2).ClipText) Or Len(RTrim$(MBox1(2).ClipText)) = 0 Or CDbl(MBox1(2).ClipText) <= 0 Then
           MsgBox "Falta Tipo de Cambio", vbInformation, MsgTitle
           Exit Sub
        End If

        If Len(RTrim$(Text2.Text)) = 0 Then
            MsgBox "Falta ingresar referencia.", vbInformation, "Sistemas"
            Text2.SetFocus
            Exit Sub
        End If
        If Ctr_Ayuda4.xclave = "" Then
          MsgBox " Ingrese codigo de empresa ", vbInformation, MsgTitle
          Ctr_Ayuda4.SetFocus
          Exit Sub
        End If
        
        VGCNx.BeginTrans
        nflag = 1
       
        If GrabarData() = 1 Then
          VGCNx.CommitTrans
          nflag = 0
          g_TipoMovi = 0
                
          If MsgBox("Desea Imprimir la Nota de Credito", vbYesNo + vbInformation, "AVISO") = vbYes Then Call ImprimirNota
                
          MBox(1) = "": MBox(2) = "": MBox(4) = "": MBox(6) = "": MBox(7) = "": MBox(8) = ""
          Combo2.ListIndex = -1
          Text2 = ""
          Text1(0).Text = Empty
          Text1(2).Text = Empty
          Text1(3).Text = Empty

          Exit Sub
        Else
           VGCNx.RollbackTrans
           nflag = 0
           g_TipoMovi = 0
           Exit Sub
        End If
       g_TipoMovi = 0
    Case 12
       g_TipoMovi = 0
       Unload Me
   End Select
   
vererror:
    If Err Then
       If nflag = 1 Then
            VGCNx.RollbackTrans
       End If
       MsgBox "Comunicarse con Sistemas ...!!" & Chr(13) & Chr(10) & Err.Number & "-" & Err.Description
       Err = 0
       Exit Sub
       Resume
    End If

End Sub

Private Sub Combo2_Click()
  Dim rs As New ADODB.Recordset
  
  If Combo2.ListCount > 0 Then
     Set rs = VGCNx.Execute("select * from vt_puntovtadocumento where puntovtacodigo='" & g_ptoventa & "' and documentocodigo='" & adll.ComboDato(Combo2.Text) & "' and empresacodigo='" & VGparametros.empresacodigo & "'")
     If rs.RecordCount > 0 Then
        MBox(1) = Format(Escadena(rs!puntovtadocserie), "0000")
        MBox(2) = Format(Escadena(rs!puntovtadoccorr), "0000000000")
     Else
        MBox(1) = Empty
        MBox(2) = Empty
     End If
     rs.Close
     
     Set rs = Nothing
  Else
     MsgBox "No tiene Serie ...Verifique!!", vbInformation, MsgTitle
     Combo2.SetFocus
  End If

End Sub

Private Sub Ctr_Ayuda3_LostFocus()
   Dim rab As New ADODB.Recordset
   Set rab = VGCNx.Execute("select * from vt_documento where documentocodigo='" & MBox(6) & "'")
   If rab.RecordCount > 0 Then
      Text2 = "" & RTrim$(Ctr_Ayuda3.xnombre) & " SEGUN DOCUMENTO " & Left$(RTrim$("" & rab!documentodescrcorta), 3) & "-" & RTrim$(MBox(7) & MBox(8))
   Else
      Text2 = "" & RTrim$(Ctr_Ayuda3.xnombre) & " SEGUN DOCUMENTO " & MBox(6) & "-" & RTrim$(MBox(7) & MBox(8))
   End If
   rab.Close
   Set rab = Nothing

End Sub

Private Sub ImprimirNota()
'FIXIT: Declare 'arrform' con un tipo de datos de enlace en tiempo de compilación          FixIT90210ae-R1672-R1B8ZE
Dim arrform(1) As Variant, arrparm(6) As Variant
Dim NombreRep As String, CadOrden As String

arrparm(0) = VGCNx.DefaultDatabase
arrparm(1) = MBox(1).Text & MBox(2).Text
arrparm(2) = CDbl(Text1(0).Text)
arrparm(3) = CDbl(Text1(2).Text)
arrparm(4) = Ctr_Ayuda4.xclave
arrparm(5) = Left$(Combo2.Text, 2)       'MBox(6).Text CODIGO DE DOCUMENTO REFERENCIA
arrform(0) = "letras='" & adll.NUMLET(numero(Round(CDbl(Text1(3)), 2))) & IIf(adll.ComboDato(Combo1.Text) = g_tiposol, "Nuevos Soles", "Dolares Americanos") & "'"

If adll.ComboDato(Combo2.Text) = "07" Then
   'NombreRep = VGparamsistem.RutaReport & "RepNotaCredito_" & Ctr_Ayuda4.xclave & ".rpt"
   NombreRep = "cc_NotaCredito_" & Ctr_Ayuda4.xclave & ".rpt"
Else
   'NombreRep = VGparamsistem.RutaReport & "RepNotaDebito_" & Ctr_Ayuda4.xclave & ".rpt"
   NombreRep = "cc_NotaDebito_" & Ctr_Ayuda4.xclave & ".rpt"
End If

Call ImpresionRptProc(NombreRep, arrform, arrparm, CadOrden, "Cuenta Corriente por Cliente")

End Sub


Private Sub MBox_KeyPress(Index As Integer, KeyAscii As Integer)
 If KeyAscii = 13 Then
         If Index Like "[1278]" Then
        If Ctr_Ayuda4.xclave = "" Then
            MsgBox "Falta seleccionar empresa", vbInformation, "Sistema"
            Ctr_Ayuda4.SetFocus
            Exit Sub
        End If

       MBox(Index) = Right$("000000000000" & RTrim$(MBox(Index).ClipText), MBox(Index).MaxLength)
       If Index = 8 Then
           If adll.VerificaDatoExistente(VGCNx, "select * from vt_cargo where documentocargo='" & MBox(6) & "' and cargonumdoc='" & RTrim$(MBox(7) & MBox(8)) & "' and clientecodigo='" & Ctr_Ayuda1.xclave & "'and empresacodigo='" & VGparametros.empresacodigo & "'") = 0 Then
              MsgBox "No existe documento del cliente...!!!", vbInformation, "AVISO"
              MBox(Index).SetFocus
 '             Exit Sub
           End If
       End If
     ElseIf Index = 4 Then
        Text1(0) = numero(MBox(4))
        If Len(RTrim$(Text1(1))) > 0 And chkInafecto.Value = 0 Then
            Text1(2) = RTrim$(numero(CDbl(Text1(0)) * CDbl(Text1(1)) / 100))
            Text1(3) = RTrim$(numero(CDbl(Text1(0)) + CDbl(Text1(2))))
        Else
            Text1(2) = "0"
            Text1(3) = RTrim$(numero(CDbl(Text1(0))))
        End If
     End If
     SendKeys "{tab}"
  End If

End Sub

Private Sub MBox_LostFocus(Index As Integer)
 If Index = 4 Then
    MBox(Index) = numero(MBox(Index))
 End If
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = 13 Then
        Text1(Index) = Text1(Index)
        If Index Like "[12]" Then
             If Len(RTrim$(Text1(1))) > 0 Then
                Text1(2) = numero(CDbl(Text1(0)) * CDbl(Text1(1)) / 100)
                Text1(3) = numero(CDbl(Text1(0)) + CDbl(Text1(2)))
             End If
        End If
        
   End If
End Sub
