VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Begin VB.Form FrmSaldos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Saldo por Cliente"
   ClientHeight    =   4890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6210
   Icon            =   "FrmSaldos.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   6210
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Filtrado por "
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
      Height          =   1575
      Left            =   120
      TabIndex        =   13
      Top             =   0
      Width           =   5895
      Begin VB.OptionButton Option2 
         Caption         =   "Vendedor"
         Height          =   255
         Left            =   360
         TabIndex        =   15
         Top             =   1200
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Cliente"
         Height          =   255
         Left            =   360
         TabIndex        =   14
         Top             =   840
         Width           =   1215
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyuCliente 
         Height          =   300
         Left            =   1785
         TabIndex        =   16
         Top             =   960
         Visible         =   0   'False
         Width           =   3915
         _ExtentX        =   6906
         _ExtentY        =   529
         XcodMaxLongitud =   0
         xcodwith        =   900
         NomTabla        =   "vt_cliente"
         ListaCampos     =   "clientecodigo(1),clienterazonsocial(1)"
         XcodCampo       =   "clientecodigo"
         XListCampo      =   "clienterazonsocial"
         ListaCamposDescrip=   "Código,Razón_Social"
         ListaCamposText =   "clientecodigo,clienterazonsocial"
         Requerido       =   0   'False
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyuVendedor 
         Height          =   315
         Left            =   1680
         TabIndex        =   18
         Top             =   960
         Visible         =   0   'False
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   556
         XcodMaxLongitud =   0
         Requerido       =   0   'False
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuempresa 
         Height          =   315
         Left            =   1680
         TabIndex        =   19
         Top             =   360
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   556
         XcodMaxLongitud =   3
         xcodwith        =   300
         NomTabla        =   "co_multiempresas"
         TituloAyuda     =   "Busqueda de Empresas"
         ListaCampos     =   "empresacodigo(1),empresadescripcion(1),agentederetencion(1)"
         XcodCampo       =   "empresacodigo"
         XListCampo      =   "empresadescripcion"
         ListaCamposDescrip=   "Codigo,Descripcion"
         ListaCamposText =   "empresacodigo,empresadescripcion,agentederetencion"
      End
      Begin VB.Label Lblempresa 
         BackStyle       =   0  'Transparent
         Caption         =   "Empresa"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   360
         TabIndex        =   20
         Top             =   360
         Width           =   975
      End
      Begin VB.Label LabelTipo 
         Caption         =   "Cliente"
         Height          =   255
         Left            =   3120
         TabIndex        =   17
         Top             =   735
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   615
      Left            =   3112
      TabIndex        =   7
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   615
      Left            =   1807
      TabIndex        =   5
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   2385
      Left            =   45
      TabIndex        =   6
      Top             =   1500
      Width           =   6030
      Begin MSComCtl2.DTPicker DTP_Fecha 
         Height          =   330
         Left            =   1665
         TabIndex        =   1
         Top             =   615
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   582
         _Version        =   393216
         Format          =   62062593
         CurrentDate     =   37579
      End
      Begin VB.ComboBox cboResumen 
         Height          =   315
         Left            =   1665
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1455
         Width           =   1785
      End
      Begin VB.ComboBox cboAcuenta 
         Height          =   315
         Left            =   1665
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1050
         Width           =   1785
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Cta 
         Height          =   375
         Left            =   1665
         TabIndex        =   0
         Top             =   195
         Width           =   4275
         _ExtentX        =   7541
         _ExtentY        =   661
         XcodMaxLongitud =   20
         xcodwith        =   700
         NomTabla        =   "ct_cuenta"
         TituloAyuda     =   "Ayuda de Cuentas"
         ListaCampos     =   "CUENTACODIGO(1),CUENTADESCRIPCION(1)"
         XcodCampo       =   "cuentacodigo"
         XListCampo      =   "cuentadescripcion"
         ListaCamposDescrip=   "Codigo,Descripcion"
         ListaCamposText =   "CUENTACODIGO,CUENTADESCRIPCION"
         Requerido       =   0   'False
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Doc 
         Height          =   300
         Left            =   1650
         TabIndex        =   4
         Top             =   1905
         Width           =   4290
         _ExtentX        =   7567
         _ExtentY        =   529
         XcodMaxLongitud =   0
         xcodwith        =   900
         NomTabla        =   "cc_tipodocumento"
         ListaCampos     =   "tdocumentocodigo(1),tdocumentodescripcion(1)"
         XcodCampo       =   "tdocumentocodigo"
         XListCampo      =   "tdocumentodescripcion"
         ListaCamposDescrip=   "Código,Descripción"
         ListaCamposText =   "tdocumentocodigo,tdocumentodescripcion"
         Requerido       =   0   'False
      End
      Begin VB.Label Label1 
         Caption         =   "Documento"
         Height          =   255
         Index           =   6
         Left            =   360
         TabIndex        =   12
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Con Resumen"
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   11
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Con Acuentas"
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   10
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Hasta la Fecha"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   9
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Cuenta Contable"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   8
         Top             =   240
         Width           =   1335
      End
   End
End
Attribute VB_Name = "FrmSaldos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adll As New dllgeneral.dll_general

Private Sub Form_Load()
   MostrarForm Me, "C2"
   Ctr_Cta.conexion VGCNx
   Ctr_Doc.conexion VGCNx
   Call CargarTipo(cboAcuenta, 3)
   Call CargarTipo(cboResumen, 3)
   Ctr_Ayuempresa.conexion VGCNx
   Ctr_AyuCliente.conexion VGCNx
   Ctr_AyuVendedor.conexion VGCNx
   
   LabelTipo.Caption = "Cliente "
   Option1.Value = True
   DTP_Fecha.Value = Format(Now, "dd/mm/yyyy")
  ' Ctr_Doc.SetFocus
 '  Ctr_Ayuempresa.SetFocus
End Sub

Private Sub cmdAceptar_Click()
If Ctr_Ayuempresa.xclave = "" Then
   MsgBox (" Ingrese codigo de empresa")
   Exit Sub
End If
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
Dim arrform(2) As Variant, arrparm(10) As Variant
Dim NombreRep As String, CadOrden As String
Dim NombrePC As String
Dim mon As String
    NombrePC = VGComputer
    arrparm(0) = VGCNx.DefaultDatabase
    arrparm(1) = NombrePC
    arrparm(2) = Format(DTP_Fecha.Value, "dd/mm/yyyy")
    arrparm(3) = cboAcuenta.ListIndex
    If Option1.Value Then
        arrparm(5) = IIf(Ctr_AyuCliente.xclave = Empty, "%", Trim(Ctr_AyuCliente.xclave))
     Else
        arrparm(5) = IIf(Ctr_AyuVendedor.xclave = Empty, "%", Trim(Ctr_AyuVendedor.xclave))
    End If
    arrparm(6) = IIf(Ctr_Cta.xclave = Empty, "%", Trim(Ctr_Cta.xclave))
    arrparm(7) = IIf(Ctr_Doc.xclave = Empty, "%", Trim(Ctr_Doc.xclave))
    arrparm(9) = Ctr_Ayuempresa.xclave
    arrparm(8) = 1
    
    arrform(0) = "@Fecha='" & Format(DTP_Fecha.Value, "dd/mm/yyyy") & "'"
    arrform(1) = "empresa='" & Ctr_Ayuempresa.xnombre & "'"
    
    NombreRep = "cc_SaldoxClienteDetalle.rpt"
    CadOrden = ""
    Call ImpresionRptProc(NombreRep, arrform, arrparm, CadOrden, "Registro de Compras ")
End Sub
Sub ImprimirConResumen()
Dim arrform(2) As Variant, arrparm(10) As Variant
Dim NombreRep As String, CadOrden As String
Dim NombreSubRep As String
Dim NombrePC As String
Dim mon As String
    NombrePC = VGComputer
    arrparm(0) = VGCNx.DefaultDatabase
    arrparm(1) = NombrePC
    arrparm(2) = Format(DTP_Fecha.Value, "dd/mm/yyyy")
    arrparm(3) = cboAcuenta.ListIndex
    arrparm(4) = IIf(Option1.Value = True, 1, 0)
    arrparm(6) = IIf(Ctr_Cta.xclave = Empty, "%", Trim(Ctr_Cta.xclave))
    arrparm(7) = IIf(Ctr_Doc.xclave = Empty, "%", Trim(Ctr_Doc.xclave))
    arrparm(8) = Ctr_Ayuempresa.xclave
    arrparm(9) = 1
    If Option1.Value Then
        arrparm(5) = IIf(Ctr_AyuCliente.xclave = Empty, "%", Trim(Ctr_AyuCliente.xclave))
     Else
        arrparm(5) = IIf(Ctr_AyuVendedor.xclave = Empty, "%", Trim(Ctr_AyuVendedor.xclave))
    End If
    arrform(0) = "@Fecha='" & Format(DTP_Fecha.Value, "dd/mm/yyyy") & "'"
    arrform(1) = "empresa='" & Ctr_Ayuempresa.xnombre & "'"
    NombreRep = "cc_SaldoxClienteDetalleResumen.rpt"
    NombreSubRep = "cc_SaldoxClienteDetalleResumen_sub.rpt"
    CadOrden = ""
    Call ImpresionRpt_SubRpt_Proc(NombreRep, arrform, arrparm, NombreSubRep, CadOrden, "Saldos de Documentos por Clientes")
End Sub

Private Sub Option1_Click()
   LabelTipo.Caption = "Cliente "
   Ctr_AyuCliente.Visible = True
   Ctr_AyuVendedor.Visible = False
 End Sub

Private Sub Option2_Click()
   LabelTipo.Caption = "Vendedor "
   Ctr_AyuCliente.Visible = False
   Ctr_AyuVendedor.Visible = True

End Sub
