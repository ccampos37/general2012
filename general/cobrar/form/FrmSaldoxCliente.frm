VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Begin VB.Form FrmSaldoxCliente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Saldo por Cliente"
   ClientHeight    =   5310
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6330
   Icon            =   "FrmSaldoxCliente.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   6330
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   495
      Left            =   3112
      TabIndex        =   9
      Top             =   4200
      Width           =   1335
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   495
      Left            =   1440
      TabIndex        =   7
      Top             =   4200
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Height          =   3945
      Left            =   45
      TabIndex        =   8
      Top             =   60
      Width           =   6030
      Begin VB.OptionButton Option1 
         Caption         =   "Solo InterEmpresas"
         Height          =   375
         Left            =   360
         TabIndex        =   19
         Top             =   3360
         Width           =   1695
      End
      Begin VB.OptionButton Option2 
         Caption         =   "No Incluye InterEmpresas"
         Height          =   375
         Left            =   2280
         TabIndex        =   18
         Top             =   3360
         Width           =   2175
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Todos "
         Height          =   375
         Left            =   4800
         TabIndex        =   17
         Top             =   3360
         Width           =   855
      End
      Begin MSComCtl2.DTPicker DTP_Fecha 
         Height          =   330
         Left            =   1665
         TabIndex        =   3
         Top             =   1575
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   582
         _Version        =   393216
         Format          =   108265473
         CurrentDate     =   37579
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuda2 
         Height          =   300
         Left            =   1665
         TabIndex        =   2
         Top             =   1185
         Width           =   4275
         _ExtentX        =   7541
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
      Begin VB.ComboBox cboResumen 
         Height          =   315
         Left            =   1665
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   2415
         Width           =   1785
      End
      Begin VB.ComboBox cboAcuenta 
         Height          =   315
         Left            =   1665
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   2010
         Width           =   1785
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Cta 
         Height          =   375
         Left            =   1665
         TabIndex        =   1
         Top             =   795
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
         TabIndex        =   6
         Top             =   2865
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
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuempresa 
         Height          =   315
         Left            =   1680
         TabIndex        =   0
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
         TabIndex        =   16
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Documento"
         Height          =   255
         Index           =   6
         Left            =   360
         TabIndex        =   15
         Top             =   2880
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Con Resumen"
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   14
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Con Acuentas"
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   13
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Hasta la Fecha"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   12
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Cliente"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   11
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Cuenta Contable"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   10
         Top             =   840
         Width           =   1335
      End
   End
End
Attribute VB_Name = "FrmSaldoxCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adll As New dllgeneral.dll_general

Private Sub Form_Load()
   MostrarForm Me, "C2"
   Ctr_Cta.conexion VGCNx
   Ctr_Ayuda2.conexion VGCNx
   Ctr_Doc.conexion VGCNx
   Call CargarTipo(cboAcuenta, 3)
   Call CargarTipo(cboResumen, 3)
   Ctr_Ayuempresa.conexion VGCNx
   Ctr_Ayuempresa.xclave = VGParametros.empresacodigo: Ctr_Ayuempresa.Ejecutar
   DTP_Fecha.Value = Format(Now, "dd/mm/yyyy")
   Option3.Value = True
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
'FIXIT: Declare 'arrform' con un tipo de datos de enlace en tiempo de compilación          FixIT90210ae-R1672-R1B8ZE
Dim arrform(2) As Variant, arrparm(10) As Variant
Dim NombreRep As String, CadOrden As String
Dim NombrePC As String
Dim mon As String
Dim ndato As String
    NombrePC = VGcomputer
    arrparm(0) = VGCNx.DefaultDatabase
    arrparm(1) = NombrePC
    arrparm(2) = Format(DTP_Fecha.Value, "dd/mm/yyyy")
    arrparm(3) = cboAcuenta.ListIndex
    arrparm(4) = "%"
    arrparm(5) = IIf(Ctr_Ayuda2.xclave = Empty, "%", RTrim$(Ctr_Ayuda2.xclave))
    arrparm(6) = IIf(Ctr_Cta.xclave = Empty, "%", RTrim$(Ctr_Cta.xclave))
    arrparm(7) = IIf(Ctr_Doc.xclave = Empty, "%", RTrim$(Ctr_Doc.xclave))
    arrparm(8) = 1
    arrparm(9) = 3
   If Option1.Value = True Then
       arrparm(9) = 1
       ndato = " SOLO INTEREMPRESAS "
     ElseIf Option2.Value = True Then
            arrparm(9) = 2
            ndato = " SIN INTEREMPRESAS "
          Else
            arrparm(9) = 3
            ndato = " TODOS "

    End If
    
    arrform(0) = "@Fecha='" & Format(DTP_Fecha.Value, "dd/mm/yyyy") & "'"
    arrform(1) = "empresa='" & Ctr_Ayuempresa.xnombre & "'"
  
    NombreRep = "cc_SaldoxClienteDetalle.rpt"
    CadOrden = ""
    Call ImpresionRptProc(NombreRep, arrform, arrparm, CadOrden, "Registro de Compras ")
End Sub
Sub ImprimirConResumen()
'FIXIT: Declare 'arrform' con un tipo de datos de enlace en tiempo de compilación          FixIT90210ae-R1672-R1B8ZE
Dim arrform(2) As Variant, arrparm(10) As Variant
Dim NombreRep As String, CadOrden As String
Dim NombreSubRep As String
Dim NombrePC As String
Dim mon As String
Dim ndato As String
    NombrePC = VGcomputer
    arrform(0) = "@Fecha='" & Format(DTP_Fecha.Value, "dd/mm/yyyy") & "'"
    arrform(1) = "tipo=' TODOS '"
    
    
    arrparm(0) = VGCNx.DefaultDatabase
    arrparm(1) = NombrePC
    arrparm(2) = Format(DTP_Fecha.Value, "dd/mm/yyyy")
    arrparm(3) = cboAcuenta.ListIndex
    arrparm(4) = "%%"
    arrparm(5) = IIf(Ctr_Ayuda2.xclave = Empty, "%%", RTrim$(Ctr_Ayuda2.xclave))
    arrparm(6) = IIf(Ctr_Cta.xclave = Empty, "%%", RTrim$(Ctr_Cta.xclave))
    arrparm(7) = IIf(Ctr_Doc.xclave = Empty, "%%", RTrim$(Ctr_Doc.xclave))
    arrparm(8) = Ctr_Ayuempresa.xclave
    arrparm(9) = 3
        If Option1.Value = True Then
       arrparm(9) = 1
       ndato = " SOLO INTEREMPRESAS "
     ElseIf Option2.Value = True Then
            arrparm(9) = 2
            ndato = " SIN INTEREMPRESAS "
          Else
            arrparm(9) = 3
            ndato = " TODOS "

    End If
    
    arrform(0) = "@Fecha='" & Format(DTP_Fecha.Value, "dd/mm/yyyy") & "'"
    arrform(1) = "empresa='" & Ctr_Ayuempresa.xnombre & "'"
    
    NombreRep = "cc_SaldoxClienteDetalleResumen.rpt"
    CadOrden = ""
    Call ImpresionRptProc(NombreRep, arrform, arrparm, CadOrden, "Saldos de Documentos por Clientes")
End Sub
