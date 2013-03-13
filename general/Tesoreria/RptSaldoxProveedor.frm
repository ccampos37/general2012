VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Begin VB.Form RptSaldoxProveedor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Saldo por Proveedor"
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6135
   Icon            =   "RptSaldoxProveedor.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   6135
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   3112
      TabIndex        =   8
      Top             =   3945
      Width           =   1215
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   1807
      TabIndex        =   7
      Top             =   3945
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   3405
      Left            =   45
      TabIndex        =   9
      Top             =   180
      Width           =   6030
      Begin MSComCtl2.DTPicker DTP_Fecha 
         Height          =   330
         Left            =   1665
         TabIndex        =   2
         Top             =   1095
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   582
         _Version        =   393216
         Format          =   23920641
         CurrentDate     =   37579
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuda2 
         Height          =   300
         Left            =   1665
         TabIndex        =   1
         Top             =   705
         Width           =   4275
         _ExtentX        =   7541
         _ExtentY        =   529
         XcodMaxLongitud =   0
         xcodwith        =   900
         NomTabla        =   "cp_proveedor"
         ListaCampos     =   "clientecodigo(1),clienterazonsocial(1)"
         XcodCampo       =   "clientecodigo"
         XListCampo      =   "clienterazonsocial"
         ListaCamposDescrip=   "C�digo,Raz�n_Social"
         ListaCamposText =   "clientecodigo,clienterazonsocial"
         Requerido       =   0   'False
      End
      Begin VB.ComboBox cboMoneda 
         Height          =   315
         Left            =   1665
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   2340
         Width           =   1785
      End
      Begin VB.ComboBox cboResumen 
         Height          =   315
         Left            =   1665
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1935
         Width           =   1785
      End
      Begin VB.ComboBox cboAcuenta 
         Height          =   315
         Left            =   1665
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1530
         Width           =   1785
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Cta 
         Height          =   375
         Left            =   1665
         TabIndex        =   0
         Top             =   315
         Width           =   4275
         _ExtentX        =   7541
         _ExtentY        =   661
         XcodMaxLongitud =   20
         xcodwith        =   700
         NomTabla        =   "cp_tipodocumento"
         TituloAyuda     =   "Ayuda de Cuentas"
         ListaCampos     =   "tdocumentocodigo(1),tdocumentodescripcion(1),tdocumentocuentasoles(1),tdocumentocuentadolares(1)"
         XcodCampo       =   "tdocumentocuentasoles"
         XListCampo      =   "tdocumentocuentadolares"
         ListaCamposDescrip=   "Codigo,Descripcion,Cuenta S/,Cuenta $"
         ListaCamposText =   "tdocumentocodigo,tdocumentodescripcion,tdocumentocuentasoles,tdocumentocuentadolares"
         Requerido       =   0   'False
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Doc 
         Height          =   300
         Left            =   1650
         TabIndex        =   6
         Top             =   2730
         Width           =   4290
         _ExtentX        =   7567
         _ExtentY        =   529
         XcodMaxLongitud =   0
         xcodwith        =   900
         NomTabla        =   "cp_tipodocumento"
         ListaCampos     =   "tdocumentocodigo(1),tdocumentodescripcion(1)"
         XcodCampo       =   "tdocumentocodigo"
         XListCampo      =   "tdocumentodescripcion"
         ListaCamposDescrip=   "C�digo,Descripci�n"
         ListaCamposText =   "tdocumentocodigo,tdocumentodescripcion"
         Requerido       =   0   'False
      End
      Begin VB.Label Label1 
         Caption         =   "Documento"
         Height          =   255
         Index           =   6
         Left            =   390
         TabIndex        =   16
         Top             =   2790
         Width           =   1185
      End
      Begin VB.Label Label1 
         Caption         =   "Moneda"
         Height          =   255
         Index           =   5
         Left            =   390
         TabIndex        =   15
         Top             =   2400
         Width           =   1185
      End
      Begin VB.Label Label1 
         Caption         =   "Con Resumen"
         Height          =   255
         Index           =   4
         Left            =   390
         TabIndex        =   14
         Top             =   2010
         Width           =   1245
      End
      Begin VB.Label Label1 
         Caption         =   "Con Acuentas"
         Height          =   255
         Index           =   3
         Left            =   390
         TabIndex        =   13
         Top             =   1590
         Width           =   1245
      End
      Begin VB.Label Label1 
         Caption         =   "Hasta la Fecha"
         Height          =   255
         Index           =   2
         Left            =   390
         TabIndex        =   12
         Top             =   1170
         Width           =   1245
      End
      Begin VB.Label Label1 
         Caption         =   "Proveedor"
         Height          =   255
         Index           =   1
         Left            =   390
         TabIndex        =   11
         Top             =   765
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Cuenta Contable"
         Height          =   255
         Index           =   0
         Left            =   390
         TabIndex        =   10
         Top             =   360
         Width           =   1245
      End
   End
End
Attribute VB_Name = "RptSaldoxProveedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adll As New dllgeneral.dll_general

Private Sub Form_Load()
   MostrarForm Me, "C2"
   Ctr_Cta.Conexion VGcnx
   Ctr_Ayuda2.Conexion VGcnx
   Ctr_Doc.Conexion VGcnx
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
Dim arrform(1) As Variant, arrparm(8) As Variant
Dim NombreRep As String, CadOrden As String
Dim NombrePC As String
Dim mon As String
    Randomize   ' Inicializa el generador de n�meros aleatorios.
    NombrePC = Trim(Str(CLng(Rnd * 10000000)))
    arrparm(0) = VGcnx.DefaultDatabase
    arrparm(1) = NombrePC
    arrparm(2) = Format(DTP_Fecha.Value, "dd/mm/yyyy")
    arrparm(3) = cboAcuenta.ListIndex
    If cboMoneda.ListIndex = 2 Then
      arrparm(4) = "%"
    Else
      arrparm(4) = Format(cboMoneda.ListIndex + 1, "00")
    End If
    arrparm(5) = IIf(Ctr_Ayuda2.xclave = Empty, "%", Trim(Ctr_Ayuda2.xclave))
    arrparm(6) = IIf(Ctr_Cta.xclave = Empty, "%", Trim(Ctr_Cta.xclave))
    arrparm(7) = IIf(Ctr_Doc.xclave = Empty, "%", Trim(Ctr_Doc.xclave))
    arrform(0) = "@Fecha='" & Format(DTP_Fecha.Value, "dd/mm/yyyy") & "'"
    NombreRep = "RepcpSaldoxProveedorDetalle.rpt"
    CadOrden = ""
    Call ImpresionRptProc(NombreRep, arrform, arrparm, CadOrden, "Registro de Compras ")
End Sub



Sub ImprimirConResumen()
Dim arrform(1) As Variant, arrparm(8) As Variant
Dim NombreRep As String, CadOrden As String
Dim NombreSubRep As String
Dim NombrePC As String
Dim mon As String
    Randomize   ' Inicializa el generador de n�meros aleatorios.
    NombrePC = Trim(Str(CLng(Rnd * 10000000)))
    arrparm(0) = VGcnx.DefaultDatabase
    arrparm(1) = NombrePC
    arrparm(2) = Format(DTP_Fecha.Value, "dd/mm/yyyy")
    arrparm(3) = cboAcuenta.ListIndex
    If cboMoneda.ListIndex = 2 Then
      arrparm(4) = "%"
    Else
      arrparm(4) = Format(cboMoneda.ListIndex + 1, "00")
    End If
    arrparm(5) = IIf(Ctr_Ayuda2.xclave = Empty, "%", Trim(Ctr_Ayuda2.xclave))
    arrparm(6) = IIf(Ctr_Cta.xclave = Empty, "%", Trim(Ctr_Cta.xclave))
    arrparm(7) = IIf(Ctr_Doc.xclave = Empty, "%", Trim(Ctr_Doc.xclave))
    arrform(0) = "@Fecha='" & Format(DTP_Fecha.Value, "dd/mm/yyyy") & "'"
    NombreRep = "RepcpSaldoxProveedorDetalleResumen.rpt"
    NombreSubRep = "RepcpSubSaldoxProveedorDetalleResumen.rpt"
    CadOrden = ""
    Call ImpresionRpt_SubRpt_Proc(NombreRep, arrform, arrparm, NombreSubRep, CadOrden, "Saldos de Documentos por Clientes")
End Sub

