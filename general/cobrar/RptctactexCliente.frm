VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Begin VB.Form RptctactexCliente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cuenta Corriente por Cliente"
   ClientHeight    =   4140
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6570
   Icon            =   "RptctactexCliente.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   6570
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   3105
      TabIndex        =   6
      Top             =   3525
      Width           =   1380
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   1485
      TabIndex        =   5
      Top             =   3525
      Width           =   1440
   End
   Begin VB.Frame Frame1 
      Height          =   3090
      Left            =   105
      TabIndex        =   3
      Top             =   120
      Width           =   6180
      Begin VB.ComboBox cboMoneda 
         Height          =   315
         Left            =   1665
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   2145
         Width           =   1605
      End
      Begin VB.ComboBox cboResumen 
         Height          =   315
         Left            =   1665
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1755
         Visible         =   0   'False
         Width           =   1605
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Cliente 
         Height          =   300
         Left            =   1665
         TabIndex        =   10
         Top             =   615
         Width           =   4365
         _ExtentX        =   7699
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
      Begin MSComCtl2.DTPicker DTP_FechaInicio 
         Height          =   330
         Left            =   1665
         TabIndex        =   11
         Top             =   1005
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         _Version        =   393216
         Format          =   56098817
         CurrentDate     =   37586
      End
      Begin MSComCtl2.DTPicker DTP_FechaFin 
         Height          =   330
         Left            =   1665
         TabIndex        =   12
         Top             =   1380
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         _Version        =   393216
         Format          =   56098817
         CurrentDate     =   37586
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Zona 
         Height          =   300
         Left            =   1665
         TabIndex        =   14
         Top             =   2595
         Visible         =   0   'False
         Width           =   4365
         _ExtentX        =   7699
         _ExtentY        =   529
         XcodMaxLongitud =   0
         xcodwith        =   900
         NomTabla        =   "vt_zona"
         ListaCampos     =   "zonacodigo(1),zonadescripcion(1)"
         XcodCampo       =   "zonacodigo"
         XListCampo      =   "zonadescripcion"
         ListaCamposDescrip=   "Código,Descripción"
         ListaCamposText =   "zonacodigo,zonadescripcion"
         Requerido       =   0   'False
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuempresa 
         Height          =   315
         Left            =   1680
         TabIndex        =   15
         Top             =   240
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
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Zona"
         Height          =   255
         Index           =   3
         Left            =   390
         TabIndex        =   13
         Top             =   2640
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.Label Label1 
         Caption         =   "Desde la Fecha"
         Height          =   255
         Index           =   0
         Left            =   390
         TabIndex        =   7
         Top             =   1065
         Width           =   1245
      End
      Begin VB.Label Label1 
         Caption         =   "Cliente"
         Height          =   255
         Index           =   1
         Left            =   390
         TabIndex        =   0
         Top             =   690
         Width           =   1245
      End
      Begin VB.Label Label1 
         Caption         =   "Hasta la Fecha"
         Height          =   255
         Index           =   2
         Left            =   390
         TabIndex        =   1
         Top             =   1440
         Width           =   1245
      End
      Begin VB.Label Label1 
         Caption         =   "Con Resumen"
         Height          =   255
         Index           =   4
         Left            =   390
         TabIndex        =   2
         Top             =   1815
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.Label Label1 
         Caption         =   "Moneda"
         Height          =   255
         Index           =   5
         Left            =   390
         TabIndex        =   4
         Top             =   2205
         Width           =   1185
      End
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
 '  Ctr_Zona.conexion VGCNx
End Sub

Private Sub cmdAceptar_Click()
If Ctr_Ayuempresa.xclave = "" Then
   MsgBox "Ingrese codigo de empresa  ", vbInformation, "AVISO"
   Exit Sub
End If
   Call Imprimir
End Sub

Private Sub cmdCancelar_Click()
   Unload Me
End Sub

Sub Imprimir()
'FIXIT: Declare 'arrform' con un tipo de datos de enlace en tiempo de compilación          FixIT90210ae-R1672-R1B8ZE
Dim arrform(2) As Variant, arrparm(8) As Variant
Dim NombreRep As String, CadOrden As String
Dim NombrePC As String
Dim mon As String
    Randomize   ' Inicializa el generador de números aleatorios.
'FIXIT: Reemplazar la función 'Str' con la función 'Str$'.                                 FixIT90210ae-R9757-R1B8ZE
    NombrePC = "##" + RTrim$(LTrim(Str(CLng(Rnd * 10000000))))
    arrparm(0) = VGCNx.DefaultDatabase
    arrparm(1) = NombrePC
    arrparm(2) = Format(DTP_FechaInicio.Value, "dd/mm/yyyy")
    arrparm(3) = Format(DTP_FechaFin.Value, "dd/mm/yyyy")
    arrparm(4) = Format(DTP_FechaInicio.Value - 1, "dd/mm/yyyy")
    If cboMoneda.ListIndex = 2 Then
      arrparm(5) = "%"
    Else
      arrparm(5) = Format(cboMoneda.ListIndex + 1, "00")
    End If
    arrparm(6) = IIf(Ctr_Cliente.xclave = Empty, "%", RTrim$(Ctr_Cliente.xclave))
    arrparm(7) = Ctr_Ayuempresa.xclave
    arrform(0) = "RangoFecha='" & "DEL " & Format(DTP_FechaInicio.Value, "dd/mm/yyyy") & " AL " & Format(DTP_FechaFin.Value, "dd/mm/yyyy") & "'"
    arrform(1) = "Empresa='" & Ctr_Ayuempresa.xnombre & "'"
    NombreRep = "cc_CtaCtexCliente.rpt"
    CadOrden = ""
    Call ImpresionRptProc(NombreRep, arrform, arrparm, CadOrden, "Cuenta Corriente por Cliente")
End Sub

