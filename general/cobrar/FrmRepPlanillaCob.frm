VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form FrmRepPlanillaCob 
   Caption         =   "Planilla de Cobranza"
   ClientHeight    =   4275
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   6990
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4275
   ScaleWidth      =   6990
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   3615
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   6225
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   0
         Left            =   1530
         TabIndex        =   2
         Top             =   2940
         Width           =   1155
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   1
         Left            =   3090
         TabIndex        =   1
         Top             =   2940
         Width           =   1155
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Vendedor 
         Height          =   405
         Left            =   1830
         TabIndex        =   3
         Top             =   1935
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   714
         XcodMaxLongitud =   3
         xcodwith        =   300
         NomTabla        =   "cp_oficina"
         TituloAyuda     =   "Ayuda de Vendedores"
         ListaCampos     =   "vendedorcodigo(1),vendedornombres(1),vendedorruc(1),vendedordireccion(1),vendedortelefono(1)"
         XcodCampo       =   "vendedorcodigo"
         XListCampo      =   "vendedornombres"
         ListaCamposDescrip=   "Codigo,Nombres,Ruc,Direccion,Telefono"
         ListaCamposText =   "vendedorcodigo,vendedornombres,vendedorruc,vendedordireccion,vendedortelefono"
         Requerido       =   0   'False
      End
      Begin MSComCtl2.DTPicker DTHasta 
         Height          =   330
         Left            =   1830
         TabIndex        =   4
         Top             =   1455
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
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
         Format          =   63897601
         CurrentDate     =   37518
      End
      Begin MSComCtl2.DTPicker DTDesde 
         Height          =   345
         Left            =   1830
         TabIndex        =   5
         Top             =   855
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   609
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
         Format          =   63897601
         CurrentDate     =   37518
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Proveedor 
         Height          =   375
         Left            =   1800
         TabIndex        =   6
         Top             =   2355
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   661
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
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuempresa 
         Height          =   375
         Left            =   1800
         TabIndex        =   7
         Top             =   360
         Width           =   3735
         _ExtentX        =   6588
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
      Begin VB.Label lbl 
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         Left            =   570
         TabIndex        =   12
         Top             =   1500
         Width           =   1080
      End
      Begin VB.Label lbl 
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   600
         TabIndex        =   11
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label lbl 
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
         Height          =   330
         Index           =   0
         Left            =   570
         TabIndex        =   10
         Top             =   1950
         Width           =   840
      End
      Begin VB.Label Label1 
         Caption         =   "Proveedor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   570
         TabIndex        =   9
         Top             =   2385
         Width           =   960
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
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   4
         Left            =   600
         TabIndex        =   8
         Top             =   480
         Width           =   975
      End
   End
   Begin Crystal.CrystalReport oCrystalReport 
      Left            =   90
      Top             =   3465
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "FrmRepPlanillaCob"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim busca As New dll_apisgen.dll_apis

Private Sub Form_Load()
   MostrarForm Me, "C2"
   DTDesde = Date
   DTHasta = Date
   Ctr_Vendedor.conexion VGCNx
   Ctr_Proveedor.conexion VGCNx
   Ctr_Ayuempresa.conexion VGCNx
 End Sub

Private Sub cmdAceptar_Click(Index As Integer)
On Error GoTo Errores
   If DTDesde > DTHasta Then
       MsgBox "Fecha Inicial debe ser mayor a Fecha Final", vbInformation, "AVISO"
       Exit Sub
   End If
   Call Imprimir
  
  Exit Sub
Errores:
  Screen.MousePointer = 1
  MsgBox "Error inesperado: " & Err.Number & "  " & Err.Description, vbExclamation
  Err = 0
  Exit Sub
End Sub

Sub Imprimir()
Dim arrform(3) As Variant, arrparm(6) As Variant
Dim NombreRep As String, CadOrden As String
Dim NombreSRep As String
Dim mon As String

    arrparm(0) = VGCNx.DefaultDatabase
    arrparm(1) = Format(DTDesde.Value, "dd/mm/yyyy")
    arrparm(2) = Format(DTHasta.Value, "dd/mm/yyyy")
    arrparm(3) = "%%"
    arrparm(4) = IIf(Ctr_Proveedor.xclave = Empty, "%%", Trim(Ctr_Proveedor.xclave))
    arrparm(5) = Ctr_Ayuempresa.xclave
    
    arrform(0) = "Desde='" & DTDesde.Value & "'"
    arrform(1) = "Hasta='" & DTHasta.Value & "'"
    arrform(2) = "Vendedor='" & Trim(Ctr_Ayuempresa.xnombre) & "'"
    CadOrden = ""
    NombreRep = "cc_PlanCobranza.rpt"
    NombreSRep = "cc_PlanCobranza_Sub.rpt"
    CadOrden = ""
    Call ImpresionRptProc(NombreRep, arrform, arrparm, CadOrden, "Planillas de Cobranzas")
End Sub


Private Sub cmdCancelar_Click(Index As Integer)
  Unload Me
End Sub

Private Sub Ctr_Vendedor_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub DTDesde_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub DTHasta_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
