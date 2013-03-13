VERSION 5.00
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "CONTROLAYUDA.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form xxxFrmNotas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de Notas de Ventas"
   ClientHeight    =   7785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7785
   ScaleWidth      =   10350
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   7275
      Left            =   120
      TabIndex        =   22
      Top             =   30
      Width           =   10035
      _ExtentX        =   17701
      _ExtentY        =   12832
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "NOTAS DE VENTAS"
      TabPicture(0)   =   "FrmNotas.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      Begin VB.Frame Frame1 
         Height          =   3915
         Left            =   150
         TabIndex        =   37
         Top             =   2970
         Width           =   9735
         Begin VB.TextBox Text2 
            Height          =   1635
            Left            =   150
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   14
            Top             =   510
            Width           =   9435
         End
         Begin VB.Frame Frame3 
            Height          =   675
            Left            =   120
            TabIndex        =   38
            Top             =   2190
            Width           =   9465
            Begin VB.TextBox Text1 
               BackColor       =   &H00C0FFFF&
               Enabled         =   0   'False
               ForeColor       =   &H00404040&
               Height          =   285
               Index           =   0
               Left            =   1410
               MaxLength       =   10
               TabIndex        =   15
               Top             =   240
               Width           =   1005
            End
            Begin VB.TextBox Text1 
               BackColor       =   &H00C0FFFF&
               ForeColor       =   &H00404040&
               Height          =   285
               Index           =   1
               Left            =   3450
               MaxLength       =   2
               TabIndex        =   16
               Top             =   210
               Width           =   675
            End
            Begin VB.TextBox Text1 
               BackColor       =   &H00C0FFFF&
               ForeColor       =   &H00404040&
               Height          =   285
               Index           =   2
               Left            =   5910
               MaxLength       =   10
               TabIndex        =   17
               Top             =   240
               Width           =   1005
            End
            Begin VB.TextBox Text1 
               BackColor       =   &H00C0FFFF&
               ForeColor       =   &H00404040&
               Height          =   285
               Index           =   3
               Left            =   8280
               MaxLength       =   10
               TabIndex        =   18
               Top             =   210
               Width           =   1005
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
               TabIndex        =   42
               Top             =   240
               Width           =   495
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
               TabIndex        =   41
               Top             =   270
               Width           =   1035
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
               TabIndex        =   40
               Top             =   270
               Width           =   1095
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
               TabIndex        =   39
               Top             =   270
               Width           =   675
            End
         End
         Begin VB.Frame Frame4 
            Height          =   930
            Left            =   4080
            TabIndex        =   47
            Top             =   2910
            Width           =   2970
            Begin VB.CommandButton cmdBotones 
               Caption         =   "&Nuevo"
               Height          =   690
               Index           =   0
               Left            =   120
               Picture         =   "FrmNotas.frx":001C
               Style           =   1  'Graphical
               TabIndex        =   48
               Top             =   180
               Width           =   870
            End
            Begin VB.CommandButton cmdBotones 
               Caption         =   "&Cancelar"
               Height          =   690
               Index           =   12
               Left            =   2040
               Picture         =   "FrmNotas.frx":045E
               Style           =   1  'Graphical
               TabIndex        =   20
               Top             =   180
               Width           =   855
            End
            Begin VB.CommandButton cmdBotones 
               Caption         =   "&Acepta"
               Height          =   690
               Index           =   11
               Left            =   1080
               Picture         =   "FrmNotas.frx":08A0
               Style           =   1  'Graphical
               TabIndex        =   19
               Top             =   180
               Width           =   870
            End
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
            TabIndex        =   46
            Top             =   210
            Width           =   9405
         End
      End
      Begin VB.Frame Frame2 
         Height          =   2505
         Left            =   150
         TabIndex        =   23
         Top             =   420
         Width           =   9735
         Begin VB.ComboBox Combo2 
            Height          =   315
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   1320
            Width           =   1245
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   8130
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   1320
            Width           =   1425
         End
         Begin VB.CommandButton cAyuda 
            Caption         =   "..."
            Height          =   285
            Left            =   3600
            TabIndex        =   12
            Top             =   2130
            Width           =   255
         End
         Begin MSMask.MaskEdBox MBox1 
            Height          =   285
            Index           =   0
            Left            =   1740
            TabIndex        =   24
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
            Left            =   5070
            TabIndex        =   25
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
            Left            =   8370
            TabIndex        =   26
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
            Left            =   1320
            TabIndex        =   0
            Top             =   930
            Width           =   8235
            _ExtentX        =   14526
            _ExtentY        =   609
            XcodMaxLongitud =   11
            xcodwith        =   800
            NomTabla        =   "vt_Cliente"
            TituloAyuda     =   "Ayuda de Clientes"
            ListaCampos     =   $"FrmNotas.frx":0CE2
            XcodCampo       =   "clientecodigo"
            XListCampo      =   "clienterazonsocial"
            ListaCamposDescrip=   "Codigo,Descripcion,Ruc,Direccion,Distrito,LimiteCred,Saldo,T,P,M,D"
            ListaCamposText =   $"FrmNotas.frx":0DC8
         End
         Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuda2 
            Height          =   315
            Left            =   6930
            TabIndex        =   8
            Top             =   1740
            Width           =   2655
            _ExtentX        =   4683
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
         Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuda3 
            Height          =   315
            Left            =   5430
            TabIndex        =   13
            Top             =   2100
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
            Height          =   255
            Index           =   3
            Left            =   6060
            TabIndex        =   4
            Top             =   1380
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
            Index           =   5
            Left            =   4530
            TabIndex        =   7
            Top             =   1770
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
            Height          =   255
            Index           =   1
            Left            =   2610
            TabIndex        =   2
            Top             =   1350
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   3
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MBox 
            Height          =   255
            Index           =   2
            Left            =   3120
            TabIndex        =   3
            Top             =   1350
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   8
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MBox 
            Height          =   255
            Index           =   4
            Left            =   1350
            TabIndex        =   6
            Top             =   1770
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   10
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MBox 
            Height          =   255
            Index           =   6
            Left            =   1350
            TabIndex        =   9
            Top             =   2130
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
            Left            =   1800
            TabIndex        =   10
            Top             =   2130
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
            Left            =   2340
            TabIndex        =   11
            Top             =   2130
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   450
            _Version        =   393216
            MaxLength       =   8
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
            Height          =   210
            Index           =   7
            Left            =   180
            TabIndex        =   45
            Top             =   1770
            Width           =   915
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
            Left            =   3540
            TabIndex        =   44
            Top             =   240
            Width           =   1395
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
            Left            =   180
            TabIndex        =   43
            Top             =   1380
            Width           =   1395
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
            Left            =   7110
            TabIndex        =   36
            Top             =   240
            Width           =   1215
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
            TabIndex        =   35
            Top             =   -300
            Width           =   1215
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000009&
            Index           =   0
            X1              =   30
            X2              =   9750
            Y1              =   600
            Y2              =   600
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00808080&
            Index           =   1
            X1              =   30
            X2              =   9720
            Y1              =   570
            Y2              =   570
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
            TabIndex        =   34
            Top             =   630
            Width           =   3795
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
            Left            =   210
            TabIndex        =   33
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label Label5 
            Caption         =   "Fecha Emision"
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
            Left            =   4740
            TabIndex        =   32
            Top             =   1380
            Width           =   1395
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
            Left            =   7380
            TabIndex        =   31
            Top             =   1380
            Width           =   915
         End
         Begin VB.Label Label5 
            Caption         =   "Fecha Vencimiento"
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
            Index           =   3
            Left            =   2850
            TabIndex        =   30
            Top             =   1770
            Width           =   1635
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
            Height          =   210
            Index           =   2
            Left            =   6000
            TabIndex        =   29
            Top             =   1770
            Width           =   1605
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
            Left            =   180
            TabIndex        =   28
            Top             =   2130
            Width           =   1305
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
            Left            =   4440
            TabIndex        =   27
            Top             =   2160
            Width           =   1305
         End
      End
   End
   Begin MSComctlLib.StatusBar Panel 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   21
      Top             =   7440
      Width           =   10350
      _ExtentX        =   18256
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
Attribute VB_Name = "xxxFrmNotas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adll As New dllgeneral.dll_general

Dim nLongicampo(6) As Integer
Dim rsdeta As New ADODB.Recordset
Dim wCabe(40)

Dim apedido As String
Dim aalmacen As String
Dim alista As String * 2



Public Function GrabarData() As Integer
    Dim j As Integer
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
    
    Dim acmd As New ADODB.Command
    Dim asql As New ADODB.Recordset
    Dim arbusca As New ADODB.Recordset

    On Error GoTo vererror
    
    GrabarData = 0
    
    '******** CABECERA DE MOVIMIENTO *****************
    
    For j = 1 To 29
        wCabe(j) = ""
    Next j
    fechasunat = Date
    apedido = MBox(6) & Trim(MBox(7) & MBox(8))
    Set asql = cn.Execute("select * from vt_pedido where pedidotipofac='" & MBox(6) & "' and pedidonrofact='" & Trim(MBox(7) & MBox(8)) & "'")
    If asql.RecordCount > 0 Then
        apedido = Escadena(asql!pedidonumero)
        wCabe(1) = Escadena(asql!puntovtacodigo)     'Escadena(asql!p)                       'Pto Venta
        wCabe(2) = Escadena(asql!pedidonumero)      'Trim(MBox(1))                       'nro pedido
        wCabe(3) = Escadena(asql!pedidonrofact)      'Trim(MBox(2))                        'nro factura
        wCabe(4) = Escadena(asql!pedidonrofact)      'Trim(MBox(3))                         'nro boleta
        wCabe(5) = Escadena(asql!pedidonrofact)      'Trim(MBox(4))                         'nro guia
        wCabe(6) = 0      'MBox(5)                       'dscto gral
        wCabe(7) = 0      'MBox(6)                       'dscto promocional
        wCabe(8) = 0      'MBox(7)                       'dscto especial
        wCabe(9) = adll.ComboDato(Combo1.Text)        'moneda
        wCabe(10) = CDbl(MBox1(2))                      'tipo de cambio
        wCabe(11) = CDbl(Escadena(asql!pedidolistaprec))    'dllgeneral.ComboDato(Combo2.Text)       'lista de precios
        wCabe(12) = " "                                'MBox(9)                      'mensajes
        wCabe(13) = Escadena(asql!modovtacodigo)     'dllgeneral.ComboDato(Combo3.Text)       'modo de venta
        wCabe(14) = MBox1(1)                         'MBox(10)                     'fecha de atencion
        wCabe(15) = Escadena(asql!formapagocodigo)     'dllgeneral.ComboDato(Combo4.Text)       'forma de pago
        wCabe(16) = Ctr_Ayuda1.xclave        ' MBox(11)                     'Proveedor
        wCabe(17) = Ctr_Ayuda2.xclave         'MBox(12)                       'vendedor
        wCabe(18) = 0    'MBox(13)                  'comision
        wCabe(19) = Escadena(asql!almacencodigo)    'Ctr_Ayuda3.xclave        'MBox(14)                     'almacen
        wCabe(20) = 0      'MBox(15)                     'otros gastos
        wCabe(21) = "0"      'MBox(16)                     'nota pedido
        wCabe(22) = "0"      'MBox(17)                     'orden de compra
        wCabe(23) = Escadena(asql!pedidoautorizacion)      'dllgeneral.ComboDato(Combo5.Text)       'autorizacion
        wCabe(24) = 0       'MBox(18)                     'dias pago
        wCabe(25) = 0                           'Total Cantidad
        wCabe(26) = Round(Text1(0), 2)                         'Total Bruto
        wCabe(27) = 0    'MBox2(8)              'total fletes --T.D.
        wCabe(28) = Round(Text1(2), 2)          'Total Igv
        wCabe(29) = Round(Text1(3), 2)         'Neto a Facturar
        wCabe(30) = Escadena(asql!pedidoentrega)    'MBox(19)                    'entrega pedido
        wCabe(31) = Escadena(asql!clienterazonsocial)  'MBox3(1)                    'nombre Proveedor
        wCabe(32) = Escadena(asql!clientedireccion)    'MBox3(3)                    'direccion
        wCabe(33) = Escadena(asql!clienteruc)  'MBox3(2)                    'ruc
        wCabe(34) = MBox(3)                  'Date                           'fechafactura
        wCabe(35) = 0                     'Total Descuentos Globales
        wCabe(36) = 0                     'Total Descuentos Cliente
        wCabe(37) = 0                     'Total Descuentos Oficina
        wCabe(38) = 0                     'Total Descuentos Item
        wCabe(39) = 0                     'Total Descuentos Linea
        wCabe(40) = 0                     'Total Descuentos x Promocion
        fechasunat = IIf(IsNull(asql!pedidofechasunat), MBox(3), asql!pedidofechasunat)
    Else
        MsgBox "Datos Incompletos del Pedido : " & apedido, vbInformation, MsgTitle
        Exit Function
    End If
    asql.Close
    
    Set asql = Nothing
    
    
    ' ** Verificando Numeracion de Documentos *****
     wCabe(2) = g_pedserie & Right("000000000000" & TraeDataSerie("select puntovtadoccorr from vt_puntovtadocumento where documentocodigo='" & g_tipoped & "' and puntovtadocserie='" & g_pedserie & "' and puntovtacodigo='" & g_ptoventa & "'", cn), 8)
     wCabe(34) = Date                       'fechafactura
     wCabe(3) = g_facserie & Right("000000000000" & TraeDataSerie("select puntovtadoccorr from vt_puntovtadocumento where documentocodigo='" & adll.ComboDato(Combo2.Text) & "' and puntovtadocserie='" & g_facserie & "' and puntovtacodigo='" & g_ptoventa & "'", cn), 8)
     wCabe(4) = adll.ComboDato(Combo2.Text)
     wCabe(5) = "0"
     If adll.VerificaDatoExistente(cn, "select * from vt_pedido where pedidonrofact='" & MBox(1) & MBox(2) & "' and pedidotipofac='" & adll.ComboDato(Combo2.Text) & "'") = 1 Then
        MsgBox "Ya existe el Documento " & adll.ComboDato(Combo2.Text) & "-" & MBox(1) & MBox(2), vbInformation, MsgTitle
        GrabarData = 0
        Exit Function
     End If
    
    '*** Verifica Serie Documentos *****
    
    Set asql = cn.Execute("select puntovtadoccorr from vt_puntovtadocumento Where documentocodigo='" & g_tipoped & "' and puntovtacodigo='" & g_ptoventa & "' and puntovtadocserie='" & MBox(1) & "'")
    If asql.RecordCount > 0 Then
       wCabe(2) = asql!puntovtadoccorr
    End If
    asql.Close
    Set asql = Nothing
    
    nsql = "Update vt_puntovtadocumento " & _
            " set puntovtadoccorr='" & Right("00000000" & Trim(CStr(wCabe(2) + 1)), 8) & "'" & _
            " Where documentocodigo='" & g_tipoped & "' and puntovtacodigo='" & g_ptoventa & "' and puntovtadocserie='" & g_pedserie & "'"
            
    cn.Execute nsql
    
    wCabe(3) = Trim(MBox(1).Text & MBox(2).Text)
     
    nsql = "Update vt_puntovtadocumento " & _
           " set puntovtadoccorr='" & Right("00000000" & Trim(CStr(wCabe(3) + 1)), 8) & "'" & _
           " Where documentocodigo='" & adll.ComboDato(Combo2.Text) & "' and puntovtacodigo='" & g_ptoventa & "' and puntovtadocserie='" & MBox(1).Text & "'"
        
    cn.Execute nsql
                   
    DoEvents
    '**cambio de documentacion
    wCabe(5) = 0
    
    DoEvents
    Set acmd.ActiveConnection = cg
    acmd.CommandType = adCmdStoredProc
    acmd.CommandText = "cp_ingresanota_pro"
    acmd.CommandTimeout = 0
    acmd.Prepared = True
    With acmd
        .Parameters("@base") = cn.DefaultDatabase
        .Parameters("@tabla") = "vt_pedido"
        .Parameters("@tipo") = IIf(adll.VerificaDatoExistente(cn, "select * from vt_pedido where pedidonumero='" & wCabe(2) & "'") = 0, "1", "2")
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
        .Parameters("@tiporefe") = MBox(6)
        .Parameters("@nrorefe") = Trim(MBox(7) & MBox(8))
        .Parameters("@fsunat") = fechasunat
    End With
    acmd.Execute
    Set acmd = Nothing
    DoEvents
        
    
   '*Grabar en los cargos ***ctacte ***
    lzona = "00"
    Set asql = cn.Execute("select * from vt_zonavendedor where vendedorcodigo='" & wCabe(17) & "'")
    If asql.RecordCount > 0 Then
        lzona = Escadena(asql!zonacodigo)
    End If
    asql.Close
    Set asql = Nothing
       
    ltipo = "1"
    If adll.VerificaDatoExistente(cn, "select * from vt_cargo where documentocargo='" & adll.ComboDato(Combo2.Text) & "' and cargonumdoc='" & Trim(MBox(1) & MBox(2)) & "'") = 0 Then
      ltipo = "1"
    Else
      ltipo = "2"
    End If
    
    If adll.VerificaDatoExistente(cn, "select * from cp_tipodocumento where tdocumentocodigo='" & adll.ComboDato(Combo2.Text) & "' and tdocumentotipo='A'") = 1 Then
      tcargo = "A"
    Else
      tcargo = "C"
    End If
        
   If wCabe(9) = g_TipoSol Then
       If tcargo = "A" Then
            cn.Execute "Update cp_proveedor " & _
                       " Set clientesaldosoles=ISNULL(clientesaldosoles,0)-" & CDbl(wCabe(29)) & _
                       "      Where clientecodigo='" & wCabe(16) & "'"
       Else
            cn.Execute "Update cp_proveedor " & _
                       " Set clientesaldosoles=ISNULL(clientesaldosoles,0)+" & CDbl(wCabe(29)) & _
                       "      Where clientecodigo='" & wCabe(16) & "'"
       
       End If
    ElseIf wCabe(9) = g_TipoDolar Then
       If tcargo = "A" Then
            cn.Execute "Update cp_proveedor " & _
                       " Set clientesaldodolares=ISNULL(clientesaldodolares,0)-" & CDbl(wCabe(29)) & _
                       "      Where clientecodigo='" & wCabe(16) & "'"
        Else
            cn.Execute "Update cp_proveedor " & _
                       " Set clientesaldodolares=ISNULL(clientesaldodolares,0)+" & CDbl(wCabe(29)) & _
                       "      Where clientecodigo='" & wCabe(16) & "'"
        End If
    End If

    Set acmd.ActiveConnection = cg
    acmd.CommandType = adCmdStoredProc
    acmd.CommandTimeout = 0
    acmd.CommandText = "cp_ingresacargovalor_pro"
    acmd.Prepared = True
    With acmd
        .Parameters("@base") = cn.DefaultDatabase
        .Parameters("@tipo") = ltipo
        .Parameters("@tabla") = "vt_cargo"
        .Parameters("@tipodocu") = adll.ComboDato(Combo2.Text)
        .Parameters("@numero") = Trim(MBox(1) & MBox(2))
        .Parameters("@cliente") = Escadena(wCabe(16))
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
        .Parameters("@referencia") = Trim(Text2)
        .Parameters("@concepto") = Escadena(Ctr_Ayuda3.xclave)
        .Parameters("@venci") = MBox(3)
    End With
    acmd.Execute
    Set acmd = Nothing
        
   MsgBox "Se Grabo Satisfactoriamente el Documento." & Chr(13) & Chr(10) & adll.ComboDato(Combo2.Text) & " >= " & MBox(1) & MBox(2), vbInformation, MsgTitle
   GrabarData = 1
    
    
vererror:
   If Err Then
      MsgBox Err.Number & "-" & Err.Description
      MsgBox "Comunicarse con Sistemas ...!!" & Chr(13) & Chr(10) & cn.Errors(0).Number & "-" & cn.Errors(0).Description
      Exit Function
   End If
End Function


Private Sub cAyuda_Click()
 nAyuda = "": nDetalle = ""
 If Len(Trim(MBox(6))) > 0 And Len(Trim(MBox(7))) > 0 And Len(Trim(MBox(8))) > 0 Then
    SendKeys "{tab}"
    Exit Sub
 End If
 
 If adll.VerificaDatoExistente(cn, "select * from vt_pedido where clientecodigo='" & Trim(Ctr_Ayuda1.xclave) & "'") = 1 Then
       Dim gfiltra(2, 2) As String
       gfiltra(1, 1) = g_tipofac: gfiltra(1, 2) = "pedidonrofact"
       gfiltra(2, 1) = g_tipobol: gfiltra(2, 2) = "pedidonroboleta"
       FrmAyuda.TipoForma = 1
       FrmAyuda.BConexion = cn
       FrmAyuda.Bdata = "0"
       FrmAyuda.BTabla = "vt_pedido"
       FrmAyuda.BCampos = "pedidotipofac as Tipo,pedidonrofact as Documento,pedidofecha as Fecha,pedidomoneda as Moneda,pedidototneto as Total"
       FrmAyuda.BOrden = "pedidofecha"
       FrmAyuda.BCondi = "clientecodigo='" & Ctr_Ayuda1.xclave & "'"
       FrmAyuda.BFiltro = gfiltra
 Else
        nAyuda = "": nDetalle = ""
        MsgBox "No existen documentos pendientes...", vbInformation, MsgTitle
        Exit Sub
 End If
 FrmAyuda.Show 1
 If Len(Escadena(nAyuda)) > 0 Then
    MBox(6) = Escadena(nAyuda): MBox(7) = Left(Escadena(nDetalle), 3): MBox(8) = Right(Escadena(nDetalle), 8)
 End If
 nAyuda = "": nDetalle = ""

End Sub



Private Sub cmdBotones_Click(Index As Integer)
   Dim asql As String
   Dim acmd As New ADODB.Command
   Dim j, nl As Integer
   
   On Error GoTo vererror
   
   Select Case Index
    Case 0
       
       MBox(1) = "": MBox(2) = "": MBox(4) = "": MBox(6) = "": MBox(7) = "": MBox(8) = ""
       Text2 = ""
       Limpiartexto Text1, 0, 3
       Ctr_Ayuda1.SetFocus
    Case 11
        If IsNull(Ctr_Ayuda1.xclave) Or Len(Trim(Ctr_Ayuda1.xclave)) = 0 Then
           MsgBox "Cliente no existe...Verifique!!!", vbInformation, MsgTitle
           Ctr_Ayuda1.SetFocus
           Exit Sub
        End If
        If IsNull(Ctr_Ayuda2.xclave) Or Len(Trim(Ctr_Ayuda2.xclave)) = 0 Then
           MsgBox "No existe Vendedor ...Verifique!!!", vbInformation, MsgTitle
           Ctr_Ayuda2.SetFocus
           Exit Sub
        End If
        If IsNull(Ctr_Ayuda3.xclave) Or Len(Trim(Ctr_Ayuda3.xclave)) = 0 Then
           MsgBox "Codigo de conceptos no existe...Verifique!!!", vbInformation, MsgTitle
           Ctr_Ayuda3.SetFocus
           Exit Sub
        End If
        If IsNull(MBox1(2).ClipText) Or Len(Trim(MBox1(2).ClipText)) = 0 Or CDbl(MBox1(2).ClipText) <= 0 Then
           MsgBox "Falta Tipo de Cambio", vbInformation, MsgTitle
           Exit Sub
        End If
'        If adll.VerificaDatoExistente(cn, "select * from cp_proveedor where clientecodigo='" & Ctr_Ayuda1.xclave & "' and clientesuspendido='1'") = 1 And Ctr_Ayuda1.xclave <> g_Eventual Then
'           MsgBox W1TXT3, vbInformation, MsgTitle
'           Exit Sub
'        End If
'        If adll.VerificaDatoExistente(cn, "select * from cp_proveedor where clientecodigo='" & Ctr_Ayuda1.xclave & "' and ((clientelimitecreddolar-clientesaldodolares)*" & MBox(8) & "+ (clientelimitecredsoles-clientesaldosoles))-" & TNeto & " <=0") = 1 And Ctr_Ayuda1.xclave <> g_Eventual Then
'           MsgBox W1TXT4, vbInformation, MsgTitle
'           Exit Sub
'        End If
'        If CDbl(MBox(4)) <> CDbl(MBox2(10)) Then
'           MsgBox "Los Totales no son iguales...Verifique!!!", vbInformation, MsgTitle
'           Exit Sub
'        End If
        
        cn.BeginTrans
        If GrabarData() = 1 Then
          cn.CommitTrans
          g_TipoMovi = 0
'          If modoventa.emitehoja = "1" Then
'             nl = IIf(modoventa.copiashoja > 0, modoventa.copiashoja, 0)
'             If nl > 0 Then
'                 For J = 1 To nl
'                    Call DocImprimir
'                 Next J
'             End If
'          End If
'          Activa 2
          Exit Sub
        Else
           cn.RollbackTrans
           g_TipoMovi = 0
'           Activa 2
           Exit Sub
        End If
       g_TipoMovi = 0
    Case 12
       g_TipoMovi = 0
       Unload Me
   End Select
   
vererror:
    If Err Then
       MsgBox "Comunicarse con Sistemas ...!!" & Chr(13) & Chr(10) & Err.Number & "-" & Err.Description
       Err = 0
       'cn.RollbackTrans
       Exit Sub
    End If

End Sub

Private Sub Combo2_Click()
  Dim rs As New ADODB.Recordset
  
  If Combo2.ListCount > 0 Then
     Set rs = cn.Execute("select * from vt_puntovtadocumento where puntovtacodigo='" & g_ptoventa & "' and documentocodigo='" & adll.ComboDato(Combo2.Text) & "'")
     If rs.RecordCount > 0 Then
        MBox(1) = Escadena(rs!puntovtadocserie)
        MBox(2) = Escadena(rs!puntovtadoccorr)
     End If
     rs.Close
     
     Set rs = Nothing
  Else
     MsgBox "No tiene Serie ...Verifique!!", vbInformation, MsgTitle
     Combo2.SetFocus
  End If

End Sub

Private Sub Form_Load()
  MostrarForm Me, "C"
  MBox1(1) = Format(Date, "DD/MM/YYYY")
    
  Call Ctr_Ayuda1.conexion(cn)
  Call Ctr_Ayuda2.conexion(cn)
  Call Ctr_Ayuda3.conexion(cn)
   
  Call adll.llenacombo(Combo2, "select * from cp_tipodocumento where tdocumentoingplan='1'", cn)
  Call adll.llenacombo(Combo1, "select * from gr_moneda", cn)
  
  MBox1(2) = Format(TraeTipoCambio(Date, cn), "##0.00")
  Text1(1) = (parametro.igv * 100)
  
End Sub

Private Sub MBox_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = 13 Then
     If Index Like "[1278]" Then
       MBox(Index) = Right("000000000000" & Trim(MBox(Index).ClipText), MBox(Index).MaxLength)
     ElseIf Index = 4 Then
        Text1(0) = Numero(MBox(4))
        If Len(Trim(Text1(1))) > 0 Then
            Text1(2) = Numero(CDbl(Text1(0)) * CDbl(Text1(1)) / 100)
            Text1(3) = Numero(CDbl(Text1(0)) + CDbl(Text1(2)))
        End If
     End If
     SendKeys "{tab}"
  End If

End Sub


Private Sub MBox_LostFocus(Index As Integer)
 If Index = 4 Then
    MBox(Index) = Numero(MBox(Index))
 End If
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = 13 Then
        Text1(Index) = Text1(Index)
        If Index Like "[12]" Then
             If Len(Trim(Text1(1))) > 0 Then
                Text1(2) = Numero(CDbl(Text1(0)) * CDbl(Text1(1)) / 100)
                Text1(3) = Numero(CDbl(Text1(0)) + CDbl(Text1(2)))
             End If
        End If
        
   End If
End Sub
