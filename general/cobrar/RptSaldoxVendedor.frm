VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Begin VB.Form RptSaldoxVendedor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Saldo por Vendedor"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7275
   Icon            =   "RptSaldoxVendedor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   7275
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   3315
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   6030
      Begin VB.ComboBox cboAcuenta 
         Height          =   315
         Left            =   1665
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1860
         Width           =   1785
      End
      Begin VB.ComboBox cboResumen 
         Height          =   315
         Left            =   1665
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   2265
         Width           =   1785
      End
      Begin VB.ComboBox cboMoneda 
         Height          =   315
         Left            =   1665
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   2670
         Width           =   1785
      End
      Begin MSComCtl2.DTPicker DTP_Fecha 
         Height          =   330
         Left            =   1665
         TabIndex        =   3
         Top             =   1425
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   582
         _Version        =   393216
         Format          =   117768193
         CurrentDate     =   37579
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Vendedor 
         Height          =   300
         Left            =   1665
         TabIndex        =   4
         Top             =   1035
         Width           =   4275
         _ExtentX        =   7541
         _ExtentY        =   529
         XcodMaxLongitud =   0
         xcodwith        =   900
         NomTabla        =   "vt_vendedor"
         ListaCampos     =   "vendedorcodigo(1),vendedornombres(1)"
         XcodCampo       =   "vendedorcodigo"
         XListCampo      =   "vendedornombres"
         ListaCamposDescrip=   "Código,Nombre"
         ListaCamposText =   "vendedorcodigo,vendedornombres"
         Requerido       =   0   'False
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Cta 
         Height          =   375
         Left            =   1665
         TabIndex        =   8
         Top             =   645
         Width           =   4275
         _ExtentX        =   7541
         _ExtentY        =   661
         XcodMaxLongitud =   20
         xcodwith        =   700
         NomTabla        =   "cc_tipodocumento"
         TituloAyuda     =   "Ayuda de Cuentas"
         ListaCampos     =   "tdocumentocodigo(1),tdocumentodescripcion(1),tdocumentocuentasoles(1),tdocumentocuentadolares(1)"
         XcodCampo       =   "tdocumentocuentasoles"
         XListCampo      =   "tdocumentocuentadolares"
         ListaCamposDescrip=   "Codigo,Descripcion,Cuenta S/,Cuenta $"
         ListaCamposText =   "tdocumentocodigo,tdocumentodescripcion,tdocumentocuentasoles,tdocumentocuentadolares"
         Requerido       =   0   'False
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyudaEmpresa 
         Height          =   315
         Left            =   1665
         TabIndex        =   15
         Top             =   240
         Width           =   4275
         _ExtentX        =   7541
         _ExtentY        =   556
         XcodMaxLongitud =   20
         xcodwith        =   200
         NomTabla        =   "co_multiempresas"
         TituloAyuda     =   "Ayuda de Empresas"
         ListaCampos     =   "empresacodigo(1),empresadescripcion(1)"
         XcodCampo       =   "empresacodigo"
         XListCampo      =   "empresadescripcion"
         ListaCamposDescrip=   "Codigo,Descripcion"
         ListaCamposText =   "empresacodigo,empresadescripcion"
         Requerido       =   0   'False
      End
      Begin VB.Label Label1 
         Caption         =   "Empresa :"
         Height          =   255
         Index           =   6
         Left            =   210
         TabIndex        =   16
         Top             =   315
         Width           =   1245
      End
      Begin VB.Label Label1 
         Caption         =   "Cuenta Contable :"
         Height          =   255
         Index           =   0
         Left            =   210
         TabIndex        =   14
         Top             =   720
         Width           =   1395
      End
      Begin VB.Label Label1 
         Caption         =   "Vendedor :"
         Height          =   255
         Index           =   1
         Left            =   210
         TabIndex        =   13
         Top             =   1125
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Hasta la Fecha :"
         Height          =   255
         Index           =   2
         Left            =   210
         TabIndex        =   12
         Top             =   1530
         Width           =   1245
      End
      Begin VB.Label Label1 
         Caption         =   "Con Acuentas :"
         Height          =   255
         Index           =   3
         Left            =   210
         TabIndex        =   11
         Top             =   1950
         Width           =   1245
      End
      Begin VB.Label Label1 
         Caption         =   "Con Resumen :"
         Height          =   255
         Index           =   4
         Left            =   210
         TabIndex        =   10
         Top             =   2370
         Width           =   1245
      End
      Begin VB.Label Label1 
         Caption         =   "Moneda :"
         Height          =   255
         Index           =   5
         Left            =   210
         TabIndex        =   9
         Top             =   2760
         Width           =   1185
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   1905
      TabIndex        =   1
      Top             =   4155
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   3210
      TabIndex        =   0
      Top             =   4155
      Width           =   1215
   End
End
Attribute VB_Name = "RptSaldoxVendedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adll As New dllgeneral.dll_general

Private Sub Form_Load()
   MostrarForm Me, "C2"
   Ctr_Cta.conexion VGCNx
   Ctr_Vendedor.conexion VGCNx
   Ctr_AyudaEmpresa.conexion VGCNx
   Call CargarTipo(cboAcuenta, 3)
   Call CargarTipo(cboResumen, 3)
   cboMoneda.Clear
   cboMoneda.AddItem g_TipoSol & "-Soles"
   cboMoneda.AddItem g_TipoDolar & "-Dolares"
   cboMoneda.AddItem "03-Ambos"
   cboMoneda.ListIndex = 2
   DTP_Fecha.Value = Format(Now, "dd/mm/yyyy")
End Sub

Private Sub cmdAceptar_Click()
  If cboResumen.ListIndex = 0 Then
     Call ImprimirConResumen
  Else
     Call ImprimirSinResumen
  End If
End Sub

Private Sub cmdCancelar_Click()
   Unload Me
End Sub

Sub ImprimirSinResumen()
'FIXIT: Declare 'arrform' con un tipo de datos de enlace en tiempo de compilación          FixIT90210ae-R1672-R1B8ZE
Dim arrform(1) As Variant, arrparm(7) As Variant
Dim NombreRep As String, CadOrden As String
Dim NombrePC As String
Dim mon As String
    Randomize   ' Inicializa el generador de números aleatorios.
'FIXIT: Reemplazar la función 'Str' con la función 'Str$'.                                 FixIT90210ae-R9757-R1B8ZE
    NombrePC = RTrim$(Str(CLng(Rnd * 10000000)))
    arrparm(0) = VGCNx.DefaultDatabase
    arrparm(1) = NombrePC
    arrparm(2) = Format(DTP_Fecha.Value, "dd/mm/yyyy")
    arrparm(3) = cboAcuenta.ListIndex
    If cboMoneda.ListIndex = 2 Then
      arrparm(4) = "%"
    Else
      arrparm(4) = Format(cboMoneda.ListIndex + 1, "00")
    End If
    arrparm(5) = IIf(Ctr_Vendedor.xclave = Empty, "%", RTrim$(Ctr_Vendedor.xclave))
    arrparm(6) = IIf(Ctr_Cta.xclave = Empty, "%", RTrim$(Ctr_Cta.xclave))
    arrform(0) = "@Fecha='" & Format(DTP_Fecha.Value, "dd/mm/yyyy") & "'"
    NombreRep = "RepccSaldoxVendedorDetalle.rpt"
    CadOrden = ""
    Call ImpresionRptProc(NombreRep, arrform, arrparm, CadOrden, "Registro de Compras ")
End Sub


Sub ImprimirConResumen()
'FIXIT: Declare 'arrform' con un tipo de datos de enlace en tiempo de compilación          FixIT90210ae-R1672-R1B8ZE
Dim arrform(1) As Variant, arrparm(7) As Variant
Dim NombreRep As String, CadOrden As String
Dim NombreSubRep As String
Dim NombrePC As String
Dim mon As String
    Randomize   ' Inicializa el generador de números aleatorios.
'FIXIT: Reemplazar la función 'Str' con la función 'Str$'.                                 FixIT90210ae-R9757-R1B8ZE
    NombrePC = RTrim$(Str(CLng(Rnd * 10000000)))
    arrparm(0) = VGCNx.DefaultDatabase
    arrparm(1) = NombrePC
    arrparm(2) = Format(DTP_Fecha.Value, "dd/mm/yyyy")
    arrparm(3) = cboAcuenta.ListIndex
    If cboMoneda.ListIndex = 2 Then
      arrparm(4) = "%"
    Else
      arrparm(4) = Format(cboMoneda.ListIndex + 1, "00")
    End If
    arrparm(5) = IIf(Ctr_Vendedor.xclave = Empty, "%", RTrim$(Ctr_Vendedor.xclave))
    arrparm(6) = IIf(Ctr_Cta.xclave = Empty, "%", RTrim$(Ctr_Cta.xclave))
    arrform(0) = "@Fecha='" & Format(DTP_Fecha.Value, "dd/mm/yyyy") & "'"
    NombreRep = "RepccSaldoxVendedorDetalleResumen.rpt"
    NombreSubRep = "RepccSubSaldoxVendedorDetalleResumen.rpt"
    CadOrden = ""
    Call ImpresionRpt_SubRpt_Proc(NombreRep, arrform, arrparm, NombreSubRep, CadOrden, "Saldos de Documentos por Clientes")
End Sub

