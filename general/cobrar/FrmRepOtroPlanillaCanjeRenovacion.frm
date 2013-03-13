VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form FrmRepOtroPlanillaCanjeRenovacion 
   Caption         =   "Planilla de Canje - Documentos Canjeados"
   ClientHeight    =   3690
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6345
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3690
   ScaleWidth      =   6345
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   3240
      Left            =   135
      TabIndex        =   0
      Top             =   210
      Width           =   6060
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   3188
         TabIndex        =   2
         Top             =   2670
         Width           =   1230
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   1703
         TabIndex        =   1
         Top             =   2670
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
         Height          =   315
         Left            =   1920
         TabIndex        =   3
         Top             =   900
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   556
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
         Format          =   117768193
         CurrentDate     =   37518
      End
      Begin MSComCtl2.DTPicker DTP_FechaInicio 
         Height          =   315
         Left            =   1935
         TabIndex        =   4
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
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
         Format          =   117768193
         CurrentDate     =   37518
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Vendedor 
         Height          =   390
         Left            =   1905
         TabIndex        =   5
         Top             =   1425
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   688
         XcodMaxLongitud =   3
         xcodwith        =   300
         NomTabla        =   "vt_vendedor"
         TituloAyuda     =   "Ayuda de Vendedores"
         ListaCampos     =   "vendedorcodigo(1),vendedornombres(1),vendedorruc(1),vendedordireccion(1),vendedortelefono(1)"
         XcodCampo       =   "vendedorcodigo"
         XListCampo      =   "vendedornombres"
         ListaCamposDescrip=   "Codigo,Nombres,Ruc,Direccion,Telefono"
         ListaCamposText =   "vendedorcodigo,vendedornombres,vendedorruc,vendedordireccion,vendedortelefono"
         Requerido       =   0   'False
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Cliente 
         Height          =   345
         Left            =   1890
         TabIndex        =   6
         Top             =   1875
         Width           =   3990
         _ExtentX        =   7038
         _ExtentY        =   609
         XcodMaxLongitud =   0
         xcodwith        =   800
         NomTabla        =   "vt_cliente"
         ListaCampos     =   "clientecodigo(1),clienterazonsocial(1)"
         XcodCampo       =   "clientecodigo"
         XListCampo      =   "clienterazonsocial"
         ListaCamposDescrip=   "Código,Descripción"
         ListaCamposText =   "clientecodigo,clienterazonsocial"
         Requerido       =   0   'False
      End
      Begin VB.Label lbl 
         Caption         =   "Cliente"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   870
         TabIndex        =   10
         Top             =   1905
         Width           =   825
      End
      Begin VB.Label lbl 
         Caption         =   "Vendedor"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   855
         TabIndex        =   9
         Top             =   1470
         Width           =   825
      End
      Begin VB.Label lbl 
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   870
         TabIndex        =   8
         Top             =   960
         Width           =   975
      End
      Begin VB.Label lbl 
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   900
         TabIndex        =   7
         Top             =   405
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
   Ctr_Vendedor.conexion cn
   Ctr_Cliente.conexion cn
End Sub

Private Sub cmdAceptar_Click()
   Call Imprimir
End Sub

Private Sub cmdCancelar_Click()
   Unload Me
End Sub

Sub Imprimir()
'FIXIT: Declare 'arrform' con un tipo de datos de enlace en tiempo de compilación          FixIT90210ae-R1672-R1B8ZE
Dim arrform(4) As Variant, arrparm(6) As Variant
Dim NombreRep As String, CadOrden As String
Dim NombreSubRep As String
Dim mon As String
   arrparm(0) = VGCNx.DefaultDatabase
   arrparm(1) = Format(DTP_FechaInicio.Value, "dd/mm/yyyy")
   arrparm(2) = Format(DTP_FechaFin.Value, "dd/mm/yyyy")
   arrparm(3) = IIf(Ctr_Vendedor.xclave = Empty, "%", RTrim$(Ctr_Vendedor.xclave))
   arrparm(4) = IIf(Ctr_Cliente.xclave = Empty, "%", RTrim$(Ctr_Cliente.xclave))
   arrparm(5) = m_Opcion
   arrform(0) = "@Titulo='" & cTitulo & "'"
   arrform(1) = "Empresa='" & g_DetalleEmpresa & "'"
   arrform(2) = "Desde='" & Format(DTP_FechaInicio.Value, "dd/mm/yyyy") & "'"
   arrform(3) = "Hasta='" & Format(DTP_FechaFin.Value, "dd/mm/yyyy") & "'"
   arrform(4) = "Vendedor='" & IIf(Ctr_Vendedor.xclave = Empty, "Todos", RTrim$(Ctr_Vendedor.xclave)) & "'"
   NombreRep = "RepccPlanOtroCanjeRenovacion.rpt"
   CadOrden = ""
   Call ImpresionRptProc(NombreRep, arrform, arrparm, CadOrden, "Planilla de " & cTitulo)
End Sub



Property Let Opcion(Valor As String)
  m_Opcion = Valor
End Property
