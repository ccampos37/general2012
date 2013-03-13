VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form FrmRepPlanillaCob_Banco 
   Caption         =   "Planillas de Cobranzas con Depositos al Banco"
   ClientHeight    =   2910
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6180
   LinkTopic       =   "Form1"
   ScaleHeight     =   2910
   ScaleWidth      =   6180
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   390
      Index           =   1
      Left            =   3270
      TabIndex        =   14
      Top             =   2340
      Width           =   1215
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   390
      Index           =   0
      Left            =   1860
      TabIndex        =   13
      Top             =   2340
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Seleccionar Filtros"
      Height          =   1845
      Left            =   195
      TabIndex        =   0
      Top             =   330
      Width           =   5700
      Begin VB.ComboBox cboTipoPago 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   2310
         Width           =   2280
      End
      Begin MSComCtl2.DTPicker DTHasta 
         Height          =   315
         Left            =   1725
         TabIndex        =   2
         Top             =   885
         Width           =   1425
         _ExtentX        =   2514
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
      Begin MSComCtl2.DTPicker DTDesde 
         Height          =   315
         Left            =   1725
         TabIndex        =   3
         Top             =   405
         Width           =   1425
         _ExtentX        =   2514
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
         Left            =   1950
         TabIndex        =   4
         Top             =   3000
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   688
         Enabled         =   0   'False
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
         Height          =   300
         Left            =   1710
         TabIndex        =   5
         Top             =   1350
         Width           =   3510
         _ExtentX        =   6191
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
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Doc 
         Height          =   300
         Left            =   1920
         TabIndex        =   6
         Top             =   3030
         Width           =   4065
         _ExtentX        =   7170
         _ExtentY        =   529
         Enabled         =   0   'False
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
      Begin VB.Label Label2 
         Caption         =   "Tipo Pago"
         Height          =   240
         Left            =   780
         TabIndex        =   12
         Top             =   2280
         Width           =   990
      End
      Begin VB.Label Label1 
         Caption         =   "Cliente"
         Height          =   270
         Index           =   0
         Left            =   495
         TabIndex        =   11
         Top             =   1395
         Width           =   945
      End
      Begin VB.Label lbl 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   2
         Left            =   495
         TabIndex        =   10
         Top             =   960
         Width           =   885
      End
      Begin VB.Label lbl 
         Caption         =   "Desde"
         Height          =   255
         Index           =   1
         Left            =   495
         TabIndex        =   9
         Top             =   450
         Width           =   855
      End
      Begin VB.Label lbl 
         Caption         =   "Vendedor"
         Height          =   315
         Index           =   0
         Left            =   720
         TabIndex        =   8
         Top             =   2940
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Documento"
         Height          =   255
         Index           =   6
         Left            =   780
         TabIndex        =   7
         Top             =   3090
         Width           =   960
      End
   End
   Begin Crystal.CrystalReport oCrystalReport 
      Left            =   180
      Top             =   3465
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "FrmRepPlanillaCob_Banco"
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
   Ctr_Vendedor.conexion cn
   Ctr_Cliente.conexion cn
   Ctr_Doc.conexion cn
   cboTipoPago.AddItem "Contado"
   cboTipoPago.AddItem "Credito"
   cboTipoPago.AddItem "Ambos"
   cboTipoPago.ListIndex = 2
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
'FIXIT: Declare 'arrform' con un tipo de datos de enlace en tiempo de compilación          FixIT90210ae-R1672-R1B8ZE
Dim arrform(0) As Variant, arrparm(7) As Variant
Dim NombreRep As String, CadOrden As String
Dim NombreSRep As String
Dim mon As String
    arrparm(0) = VGCNx.DefaultDatabase
    arrparm(1) = Format(DTDesde.Value, "dd/mm/yyyy")
    arrparm(2) = Format(DTHasta.Value, "dd/mm/yyyy")
    arrparm(3) = IIf(Ctr_Vendedor.xclave = Empty, "%", RTrim$(Ctr_Vendedor.xclave))
    arrparm(4) = IIf(Ctr_Cliente.xclave = Empty, "%", RTrim$(Ctr_Cliente.xclave))
    If cboTipoPago.ListIndex = 2 Then
        arrparm(5) = "%"
    ElseIf cboTipoPago.ListIndex = 0 Then
        arrparm(5) = "CO"
    Else
        arrparm(5) = "CR"
    End If
    arrparm(6) = IIf(Ctr_Doc.xclave = Empty, "%", RTrim$(Ctr_Doc.xclave))
    NombreRep = "RepccPlanCobranza_banco.rpt"
    'NombreSRep = "RepccSubPlanCobranza.rpt"
    CadOrden = ""
    Call ImpresionRpt_SubRpt_Proc(NombreRep, arrform, arrparm, NombreSRep, CadOrden, "Planilla de Cobranza")
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


