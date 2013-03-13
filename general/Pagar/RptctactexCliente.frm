VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Begin VB.Form RptctactexCliente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cuenta Corriente por Proveedor"
   ClientHeight    =   3645
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6600
   Icon            =   "RptctactexCliente.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   6600
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2865
      Left            =   210
      TabIndex        =   2
      Top             =   210
      Width           =   6180
      Begin VB.ComboBox cboResumen 
         Height          =   315
         Left            =   1665
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1875
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.ComboBox cboMoneda 
         Height          =   315
         Left            =   1665
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   2265
         Width           =   1245
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Cliente 
         Height          =   300
         Left            =   1665
         TabIndex        =   5
         Top             =   735
         Width           =   4365
         _ExtentX        =   7699
         _ExtentY        =   529
         XcodMaxLongitud =   0
         xcodwith        =   900
         NomTabla        =   "cp_proveedor"
         ListaCampos     =   "clientecodigo(1),clienterazonsocial(1)"
         XcodCampo       =   "clientecodigo"
         XListCampo      =   "clienterazonsocial"
         ListaCamposDescrip=   "Código,Razón_Social"
         ListaCamposText =   "clientecodigo,clienterazonsocial"
         Requerido       =   0   'False
      End
      Begin MSComCtl2.DTPicker DTP_FechaInicio 
         Height          =   330
         Left            =   1665
         TabIndex        =   6
         Top             =   1125
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         _Version        =   393216
         Format          =   27262977
         CurrentDate     =   37586
      End
      Begin MSComCtl2.DTPicker DTP_FechaFin 
         Height          =   330
         Left            =   1665
         TabIndex        =   7
         Top             =   1500
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         _Version        =   393216
         Format          =   27262977
         CurrentDate     =   37586
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuempresa 
         Height          =   375
         Left            =   1680
         TabIndex        =   13
         Top             =   240
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   661
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
      Begin VB.Label Label1 
         Caption         =   "Empresa"
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   14
         Top             =   360
         Width           =   1245
      End
      Begin VB.Label Label1 
         Caption         =   "Moneda"
         Height          =   255
         Index           =   5
         Left            =   390
         TabIndex        =   12
         Top             =   2325
         Width           =   1185
      End
      Begin VB.Label Label1 
         Caption         =   "Con Resumen"
         Height          =   255
         Index           =   4
         Left            =   390
         TabIndex        =   11
         Top             =   1935
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.Label Label1 
         Caption         =   "Hasta la Fecha"
         Height          =   255
         Index           =   2
         Left            =   390
         TabIndex        =   10
         Top             =   1560
         Width           =   1245
      End
      Begin VB.Label Label1 
         Caption         =   "Proveedor"
         Height          =   255
         Index           =   1
         Left            =   390
         TabIndex        =   9
         Top             =   810
         Width           =   1245
      End
      Begin VB.Label Label1 
         Caption         =   "Desde la Fecha"
         Height          =   255
         Index           =   0
         Left            =   390
         TabIndex        =   8
         Top             =   1185
         Width           =   1245
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   1590
      TabIndex        =   1
      Top             =   3120
      Width           =   1440
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   3210
      TabIndex        =   0
      Top             =   3135
      Width           =   1380
   End
End
Attribute VB_Name = "RptctactexCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adll As New dllgeneral.dll_general

Private Sub Form_Load()
   MostrarForm Me, "C2"
   Call CargarTipo(cboResumen, 3)
   cboMoneda.Clear
   cboMoneda.AddItem g_TipoSol & "-Soles"
   cboMoneda.AddItem g_TipoDolar & "-Dolares"
   cboMoneda.AddItem "03-Ambos"
   cboMoneda.ListIndex = 2
   DTP_FechaInicio.Value = "01/" & Format(Month(Now), "00") & "/" & Year(Date)
   DTP_FechaFin.Value = Format(Date, "dd/mm/yyyy")
   Ctr_Cliente.conexion VGCNx
   Ctr_Ayuempresa.conexion VGCNx
   If VGparametros.sistemamultiempresas = False Then
      Ctr_Ayuempresa.xclave = VGparametros.empresacodigo: Ctr_Ayuempresa.Ejecutar
      Ctr_Ayuempresa.Enabled = False
   End If
End Sub

Private Sub cmdAceptar_Click()
   Call Imprimir
End Sub

Private Sub cmdCancelar_Click()
   Unload Me
End Sub

Sub Imprimir()
Dim arrform(2) As Variant, arrparm(8) As Variant
Dim NombreRep As String, CadOrden As String
Dim mon As String
    Randomize   ' Inicializa el generador de números aleatorios.
    arrparm(0) = VGCNx.DefaultDatabase
    arrparm(1) = VGComputer
    arrparm(2) = Format(DTP_FechaInicio.Value, "dd/mm/yyyy")
    arrparm(3) = Format(DTP_FechaFin.Value, "dd/mm/yyyy")
    arrparm(4) = Format(DTP_FechaInicio.Value - 1, "dd/mm/yyyy")
    If cboMoneda.ListIndex = 2 Then
      arrparm(5) = "%%"
    Else
      arrparm(5) = Format(cboMoneda.ListIndex + 1, "00")
    End If
    arrparm(6) = IIf(Ctr_Cliente.xclave = Empty, "%%", Trim$(Ctr_Cliente.xclave))
    arrparm(7) = IIf(Ctr_Ayuempresa.xclave = Empty, "%%", Trim$(Ctr_Ayuempresa.xclave))
    
    NombreRep = "cp_CtaCtexProveedor.rpt"
    arrform(0) = "RangoFecha='" & "DEL " & Format(DTP_FechaInicio.Value, "dd/mm/yyyy") & " AL " & Format(DTP_FechaFin.Value, "dd/mm/yyyy") & "'"
    arrform(1) = "Empresa='" & IIf(Ctr_Ayuempresa.xclave = Empty, "EMPRESA : TODAS ", Trim$(Ctr_Ayuempresa.xnombre)) & "'"
    CadOrden = ""
    Call ImpresionRptProc(NombreRep, arrform, arrparm, CadOrden, "Cuenta Corriente de Cuentas por Pagar")
End Sub
