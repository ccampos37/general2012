VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form FrmRepOtroPlanillaCanjeRenovacion 
   Caption         =   "Planilla de Canje - Documentos Canjeados"
   ClientHeight    =   4110
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6345
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4110
   ScaleWidth      =   6345
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   3855
      Left            =   135
      TabIndex        =   0
      Top             =   150
      Width           =   6075
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   3188
         TabIndex        =   2
         Top             =   3150
         Width           =   1230
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   1703
         TabIndex        =   1
         Top             =   3150
         Width           =   1230
      End
      Begin Crystal.CrystalReport oCrystalReport 
         Left            =   135
         Top             =   2640
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin MSComCtl2.DTPicker DTP_FechaFin 
         Height          =   375
         Left            =   1920
         TabIndex        =   3
         Top             =   1440
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   27262977
         CurrentDate     =   37518
      End
      Begin MSComCtl2.DTPicker DTP_FechaInicio 
         Height          =   375
         Left            =   1920
         TabIndex        =   4
         Top             =   840
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   27262977
         CurrentDate     =   37518
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Cliente 
         Height          =   375
         Left            =   1920
         TabIndex        =   5
         Top             =   2400
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   661
         XcodMaxLongitud =   0
         xcodwith        =   800
         NomTabla        =   "cp_proveedor"
         ListaCampos     =   "clientecodigo(1),clienterazonsocial(1)"
         XcodCampo       =   "clientecodigo"
         XListCampo      =   "clienterazonsocial"
         ListaCamposDescrip=   "Código,Descripción"
         ListaCamposText =   "clientecodigo,clienterazonsocial"
         Requerido       =   0   'False
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyuEmpresa 
         Height          =   375
         Left            =   1920
         TabIndex        =   9
         Top             =   360
         Width           =   3855
         _ExtentX        =   6800
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
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuoficina 
         Height          =   375
         Left            =   1920
         TabIndex        =   11
         Top             =   1800
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   661
         XcodMaxLongitud =   3
         xcodwith        =   200
         NomTabla        =   "cp_oficina"
         ListaCampos     =   "vendedorcodigo(1),vendedornombres(1)"
         XcodCampo       =   "vendedorcodigo"
         XListCampo      =   "vendedornombres"
         ListaCamposDescrip=   "Codigo,Descripcion"
         ListaCamposText =   "vendedorcodigo,vendedornombres"
         Requerido       =   0   'False
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Oficina"
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
         Height          =   225
         Index           =   2
         Left            =   840
         TabIndex        =   12
         Top             =   1950
         Width           =   765
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Empresa"
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
         Height          =   225
         Index           =   3
         Left            =   840
         TabIndex        =   10
         Top             =   480
         Width           =   1005
      End
      Begin VB.Label lbl 
         Caption         =   "Proveedor"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   3
         Left            =   840
         TabIndex        =   8
         Top             =   2385
         Width           =   915
      End
      Begin VB.Label lbl 
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   2
         Left            =   870
         TabIndex        =   7
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label lbl 
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   1
         Left            =   900
         TabIndex        =   6
         Top             =   885
         Width           =   855
      End
   End
End
Attribute VB_Name = "FrmRepOtroPlanillaCanjeRenovacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim m_Opcion As String
Dim adll As New dllgeneral.dll_general
Dim cTitulo As String

Private Sub Form_Load()
   MostrarForm Me, "C2"
   DTP_FechaInicio.Value = "01/" & Format(Month(Now), "00") & "/" & Year(Date)
   DTP_FechaFin.Value = Format(Date, "dd/mm/yyyy")
   Select Case m_Opcion
      Case "1": cTitulo = "Canje"
      Case "2": cTitulo = "Renovación"
   End Select
   Me.Caption = "Planilla de " & cTitulo
   Call Ctr_Ayuempresa.conexion(VGCNx)
   Call Ctr_Ayuoficina.conexion(VGCNx)
   Call Ctr_Cliente.conexion(VGCNx)
End Sub

Private Sub cmdAceptar_Click()
   Call Imprimir
End Sub

Private Sub cmdCancelar_Click()
   Unload Me
End Sub

Sub Imprimir()
Dim arrform(4) As Variant
Dim arrparm(7) As Variant
Dim NombreRep As String, CadOrden As String
Dim NombreSubRep As String
Dim mon As String
   arrparm(0) = VGCNx.DefaultDatabase
   arrparm(1) = DTP_FechaInicio.Value
   arrparm(2) = DTP_FechaFin.Value
   arrparm(3) = IIf(Ctr_Cliente.xclave = Empty, "%%", Trim$(Ctr_Cliente.xclave))
   arrparm(4) = m_Opcion
   arrparm(5) = "%%"
   arrparm(6) = Ctr_Ayuempresa.xclave
   
   
   
   arrform(0) = "@Titulo='" & cTitulo & "'"
   arrform(1) = "Desde='" & Format(DTP_FechaInicio.Value, "dd/mm/yyyy") & "'"
   arrform(2) = "Hasta='" & Format(DTP_FechaFin.Value, "dd/mm/yyyy") & "'"
   arrform(3) = "Vendedor='" & IIf(Ctr_Ayuoficina.xclave = Empty, "Todos", Trim$(Ctr_Ayuoficina.xclave)) & "'"
   NombreRep = "cp_PlanCanjeRenovacion.rpt"
   CadOrden = ""
   Call ImpresionRptProc(NombreRep, arrform, arrparm, CadOrden, "Planilla de " & cTitulo)
End Sub


Property Let Opcion(valor As String)
  m_Opcion = valor
End Property
