VERSION 5.00
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "CONTROLAYUDA.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form xxxFrmAnulaNota 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Anulación de Notas de Ventas"
   ClientHeight    =   7860
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10320
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7860
   ScaleWidth      =   10320
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   7275
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   10035
      _ExtentX        =   17701
      _ExtentY        =   12832
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "NOTAS DE VENTAS"
      TabPicture(0)   =   "FrmAnulaNota.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      Begin VB.Frame Frame2 
         Height          =   2505
         Left            =   150
         TabIndex        =   25
         Top             =   420
         Width           =   9735
         Begin VB.CommandButton cAyuda2 
            Caption         =   "..."
            Height          =   285
            Left            =   4410
            TabIndex        =   48
            Top             =   1350
            Width           =   255
         End
         Begin VB.CommandButton cAyuda 
            Caption         =   "..."
            Height          =   285
            Left            =   3600
            TabIndex        =   26
            Top             =   2130
            Width           =   255
         End
         Begin VB.ComboBox Combo1 
            Enabled         =   0   'False
            Height          =   315
            Left            =   8130
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   1320
            Width           =   1425
         End
         Begin VB.ComboBox Combo2 
            Height          =   315
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   1320
            Width           =   1245
         End
         Begin MSMask.MaskEdBox MBox1 
            Height          =   285
            Index           =   0
            Left            =   1740
            TabIndex        =   27
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
            TabIndex        =   28
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
            TabIndex        =   29
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
            ListaCampos     =   $"FrmAnulaNota.frx":001C
            XcodCampo       =   "clientecodigo"
            XListCampo      =   "clienterazonsocial"
            ListaCamposDescrip=   "Codigo,Descripcion,Ruc,Direccion,Distrito,LimiteCred,Saldo,T,P,M,D"
            ListaCamposText =   $"FrmAnulaNota.frx":0102
         End
         Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuda2 
            Height          =   315
            Left            =   6930
            TabIndex        =   10
            Top             =   1740
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   556
            Enabled         =   0   'False
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
            TabIndex        =   30
            Top             =   2100
            Width           =   4215
            _ExtentX        =   7435
            _ExtentY        =   556
            Enabled         =   0   'False
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
            Left            =   6090
            TabIndex        =   6
            Top             =   1320
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   529
            _Version        =   393216
            ClipMode        =   1
            AllowPrompt     =   -1  'True
            Enabled         =   0   'False
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
            Height          =   300
            Index           =   5
            Left            =   4530
            TabIndex        =   9
            Top             =   1755
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   529
            _Version        =   393216
            ClipMode        =   1
            AllowPrompt     =   -1  'True
            Enabled         =   0   'False
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
            TabIndex        =   8
            Top             =   1770
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            _Version        =   393216
            Enabled         =   0   'False
            MaxLength       =   10
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MBox 
            Height          =   255
            Index           =   6
            Left            =   1350
            TabIndex        =   31
            Top             =   2130
            Width           =   435
            _ExtentX        =   767
            _ExtentY        =   450
            _Version        =   393216
            Enabled         =   0   'False
            MaxLength       =   3
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MBox 
            Height          =   255
            Index           =   7
            Left            =   1800
            TabIndex        =   32
            Top             =   2130
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   450
            _Version        =   393216
            Enabled         =   0   'False
            MaxLength       =   3
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MBox 
            Height          =   255
            Index           =   8
            Left            =   2340
            TabIndex        =   33
            Top             =   2130
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   450
            _Version        =   393216
            Enabled         =   0   'False
            MaxLength       =   8
            PromptChar      =   "_"
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
            TabIndex        =   46
            Top             =   2160
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
            Left            =   180
            TabIndex        =   45
            Top             =   2130
            Width           =   1305
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
            TabIndex        =   44
            Top             =   1770
            Width           =   1605
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
            TabIndex        =   43
            Top             =   1800
            Width           =   1665
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
            TabIndex        =   42
            Top             =   1380
            Width           =   915
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
            TabIndex        =   41
            Top             =   1350
            Width           =   1395
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
            TabIndex        =   40
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
            TabIndex        =   39
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
            TabIndex        =   38
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
            Left            =   7110
            TabIndex        =   37
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
            Left            =   180
            TabIndex        =   36
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
            Left            =   3540
            TabIndex        =   35
            Top             =   240
            Width           =   1395
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
            TabIndex        =   34
            Top             =   1770
            Width           =   915
         End
      End
      Begin VB.Frame Frame1 
         Height          =   3915
         Left            =   150
         TabIndex        =   12
         Top             =   2970
         Width           =   9735
         Begin VB.Frame Frame4 
            Height          =   930
            Left            =   4080
            TabIndex        =   23
            Top             =   2910
            Width           =   1980
            Begin VB.CommandButton cmdBotones 
               Caption         =   "&Acepta"
               Height          =   690
               Index           =   11
               Left            =   90
               Picture         =   "FrmAnulaNota.frx":01C7
               Style           =   1  'Graphical
               TabIndex        =   4
               Top             =   180
               Width           =   870
            End
            Begin VB.CommandButton cmdBotones 
               Caption         =   "&Cancelar"
               Height          =   690
               Index           =   12
               Left            =   1050
               Picture         =   "FrmAnulaNota.frx":0609
               Style           =   1  'Graphical
               TabIndex        =   5
               Top             =   180
               Width           =   855
            End
         End
         Begin VB.Frame Frame3 
            Height          =   675
            Left            =   120
            TabIndex        =   14
            Top             =   2190
            Width           =   9465
            Begin VB.TextBox Text1 
               BackColor       =   &H00C0FFFF&
               Enabled         =   0   'False
               ForeColor       =   &H00404040&
               Height          =   285
               Index           =   3
               Left            =   8280
               MaxLength       =   10
               TabIndex        =   18
               Top             =   210
               Width           =   1005
            End
            Begin VB.TextBox Text1 
               BackColor       =   &H00C0FFFF&
               Enabled         =   0   'False
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
               Enabled         =   0   'False
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
               TabIndex        =   22
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
               TabIndex        =   21
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
               TabIndex        =   20
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
               TabIndex        =   19
               Top             =   240
               Width           =   495
            End
         End
         Begin VB.TextBox Text2 
            Enabled         =   0   'False
            Height          =   1635
            Left            =   150
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   13
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
            TabIndex        =   24
            Top             =   210
            Width           =   9405
         End
      End
   End
   Begin MSComctlLib.StatusBar Panel 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   47
      Top             =   7515
      Width           =   10320
      _ExtentX        =   18203
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
Attribute VB_Name = "xxxFrmAnulaNota"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adll As New dllgeneral.dll_general

Dim nLongicampo(6) As Integer
Dim rsdeta As New ADODB.Recordset

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
    
    If adll.VerificaDatoExistente(cn, "select * from vt_abono where documentoabono='" & adll.ComboDato(Combo2) & "' and  abononumdoc='" & Trim(MBox(1) & MBox(2)) & "'") = 0 Then
         cn.Execute " Update vt_pedido " & _
                     " Set pedidofechaanu='" & Date & "'," & _
                     "     pedidocondicionfactura='1' " & _
                     " where pedidotipofac='" & adll.ComboDato(Combo2) & "' and pedidonrofact='" & Trim(MBox(1) & MBox(2)) & "'"
                     
         cn.Execute " Update vt_cargo " & _
                    " Set cargoapeflgreg='1' " & _
                    " where documentocargo='" & adll.ComboDato(Combo2) & "' and cargonumdoc='" & Trim(MBox(1) & MBox(2)) & "'"
                    
        If adll.ComboDato(Combo1.Text) = g_TipoSol Then
            If adll.VerificaDatoExistente(cn, "select * from cp_tipodocumento where tdocumentocodigo='" & adll.ComboDato(Combo2) & "' and tdocumentotipo='A'") = 1 Then
                cn.Execute " Update cp_proveedor " & _
                           " Set clientesaldosoles=isnull(clientesaldosoles,0)+" & CDbl(Text1(3)) & _
                           " Where clientecodigo='" & Ctr_Ayuda1.xclave & "'"
            Else
                cn.Execute " Update cp_proveedor " & _
                           " Set clientesaldosoles=isnull(clientesaldosoles,0)-" & CDbl(Text1(3)) & _
                           " Where clientecodigo='" & Ctr_Ayuda1.xclave & "'"
            
            End If
        ElseIf adll.ComboDato(Combo1.Text) = g_TipoDolar Then
            If adll.VerificaDatoExistente(cn, "select * from cp_tipodocumento where tdocumentocodigo='" & adll.ComboDato(Combo2) & "' and tdocumentotipo='A'") = 1 Then
                cn.Execute " Update cp_proveedor " & _
                           " Set clientesaldodolares=isnull(clientesaldodolares,0)+" & CDbl(Text1(3)) & _
                           " Where clientecodigo='" & Ctr_Ayuda1.xclave & "'"
            Else
                cn.Execute " Update cp_proveedor " & _
                           " Set clientesaldodolares=isnull(clientesaldodolares,0)-" & CDbl(Text1(3)) & _
                           " Where clientecodigo='" & Ctr_Ayuda1.xclave & "'"
            End If
            
        End If
        MsgBox "Se Anulo Satisfactoriamente el Documento." & Chr(13) & Chr(10) & adll.ComboDato(Combo2.Text) & " >= " & MBox(1) & MBox(2), vbInformation, MsgTitle
        GrabarData = 1
    Else
        MsgBox "No se puede anular el documento tiene abonos." & Chr(13) & Chr(10) & adll.ComboDato(Combo2.Text) & " >= " & MBox(1) & MBox(2), vbInformation, MsgTitle
        GrabarData = 0
    End If
    
    
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



Private Sub cAyuda2_Click()
 nAyuda = "": nDetalle = ""
 If Len(Trim(MBox(1))) > 0 And Len(Trim(MBox(2))) > 0 Then
    SendKeys "{tab}"
    Exit Sub
 End If
 
 If adll.VerificaDatoExistente(cn, "select * from vt_cargo where documentocargo='" & adll.ComboDato(Combo2) & "'") = 1 Then
       Dim sfiltra(1, 2) As String
       sfiltra(1, 1) = "Documento": sfiltra(1, 2) = "cargonumdoc"
       FrmAyuda.TipoForma = 1
       FrmAyuda.BConexion = cn
       FrmAyuda.Bdata = "0"
       FrmAyuda.BTabla = "vt_cargo"
       FrmAyuda.BCampos = "documentocargo as Tipo,cargonumdoc as Documento,cargoapefecemi as Fecha,monedacodigo as Moneda,cargoapeimpape as Total"
       FrmAyuda.BOrden = "cargoapefecemi"
       FrmAyuda.BCondi = "documentocargo='" & adll.ComboDato(Combo2) & "'"
       FrmAyuda.BFiltro = sfiltra
 Else
        nAyuda = "": nDetalle = ""
        MsgBox "No existen documentos pendientes...", vbInformation, MsgTitle
        Exit Sub
 End If
 FrmAyuda.Show 1
 If Len(Escadena(nAyuda)) > 0 Then
    MBox(1) = Left(Escadena(nDetalle), 3): MBox(2) = Right(Escadena(nDetalle), 8)
    Call MBox_KeyPress(2, 13)
 End If
 nAyuda = "": nDetalle = ""

End Sub

Private Sub cmdBotones_Click(Index As Integer)
   Dim asql As String
   Dim acmd As New ADODB.Command
   Dim j, nl As Integer
   
   On Error GoTo vererror
   
   Select Case Index
    Case 11
        If MsgBox("Desea Anular el Documento?", vbInformation, MsgTitle) = vbNo Then
            Exit Sub
        End If

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
       Exit Sub
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
  Dim rb As New ADODB.Recordset
  Dim rb2 As New ADODB.Recordset
  
  If KeyAscii = 13 Then
     If Index Like "[1278]" Then
        MBox(Index) = Right("000000000000" & Trim(MBox(Index).ClipText), MBox(Index).MaxLength)
        If Index = 2 Then
            Set rb = cn.Execute("select * from vt_cargo where documentocargo='" & adll.ComboDato(Combo2.Text) & "' and cargonumdoc='" & Trim(MBox(1) & MBox(2)) & "' and clientecodigo='" & Ctr_Ayuda1.xclave & "'")
            If rb.RecordCount > 0 Then
                MBox1(1) = Format(rb!fechaact, "dd/mm/yyyy")
                MBox1(2) = Format(rb!cargoapetipcam, "##0.00")
                
                MBox(3) = Format(rb!cargoapefecemi, "dd/mm/yyyy")
                Combo1.ListIndex = VerificaCombo(Combo1, Escadena(rb!monedacodigo))
                MBox(5) = Format(rb!cargoapefecvct, "dd/mm/yyyy")
                
                Ctr_Ayuda2.xclave = Escadena(rb!vendedorcodigo)
                Ctr_Ayuda2.Ejecutar
                
                Set rb2 = cn.Execute("select * from vt_pedido where pedidotipofac='" & rb!documentocargo & "' and pedidonrofact='" & rb!cargonumdoc & "' and clientecodigo='" & rb!clientecodigo & "'")
                If rb2.RecordCount > 0 Then
                    MBox(4) = Numero(rb2!pedidototbruto)
                    MBox(6) = Escadena(rb2!pedidotiporefe)
                    MBox(7) = Left(Escadena(rb2!pedidonrorefe), 3)
                    MBox(8) = Right(Escadena(rb2!pedidonrorefe), 8)
                    Text1(0) = Numero(rb2!pedidototbruto)
                    If rb2!pedidototimpuesto > 0 Then
                        Text1(1) = Format((rb2!pedidototimpuesto * 100) / rb2!pedidototbruto, "##0.00")
                        Text1(2) = Numero(rb2!pedidototimpuesto)
                    Else
                        Text1(1) = Numero(0)
                        Text1(2) = Numero(rb2!pedidototimpuesto)
                    End If
                    Text1(3) = Numero(rb2!pedidototneto)
                End If
                rb2.Close
                Ctr_Ayuda3.xclave = Escadena(rb!conceptocodigo)
                Ctr_Ayuda3.Ejecutar
                
                Text2 = Escadena(Trim(rb!cargoaperefere))
         
            End If
            rb.Close
       End If
     ElseIf Index = 4 Then
        Text1(0) = Numero(MBox(4))
        If Len(Trim(Text1(1))) > 0 Then
            Text1(2) = Numero(CDbl(Text1(0)) * CDbl(Text1(1)) / 100)
            Text1(3) = Numero(CDbl(Text1(0)) + CDbl(Text1(2)))
        End If
     End If
     SendKeys "{tab}"
  End If
  
  Set rb2 = Nothing
  Set rb = Nothing

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


