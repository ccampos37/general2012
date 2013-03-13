VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmRepListadoDocumentos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Documentos"
   ClientHeight    =   3705
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6285
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   6285
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Seleccionar Criterios"
      Height          =   2610
      Left            =   75
      TabIndex        =   6
      Top             =   90
      Width           =   6105
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_TipoDoc 
         Height          =   300
         Left            =   1275
         TabIndex        =   3
         Top             =   1755
         Width           =   3180
         _ExtentX        =   5609
         _ExtentY        =   529
         XcodMaxLongitud =   0
         xcodwith        =   350
         NomTabla        =   "cc_tipodocumento"
         ListaCampos     =   "tdocumentocodigo(1),tdocumentodescripcion(1)"
         XcodCampo       =   "tdocumentocodigo"
         XListCampo      =   "tdocumentodescripcion"
         ListaCamposDescrip=   "Código,Descripción"
         ListaCamposText =   "tdocumentocodigo,tdocumentodescripcion"
         Requerido       =   0   'False
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Cliente 
         Height          =   345
         Left            =   1275
         TabIndex        =   2
         Top             =   1305
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   609
         XcodMaxLongitud =   0
         xcodwith        =   900
         NomTabla        =   "vt_cliente"
         ListaCampos     =   "clientecodigo(1),clienterazonsocial(1)"
         XcodCampo       =   "clientecodigo"
         XListCampo      =   "clienterazonsocial"
         ListaCamposDescrip=   "Código,Descripción"
         ListaCamposText =   "clientecodigo,clienterazonsocial"
         Requerido       =   0   'False
      End
      Begin MSComCtl2.DTPicker DTHasta 
         Height          =   330
         Left            =   1290
         TabIndex        =   1
         Top             =   840
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
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
         Height          =   330
         Left            =   1290
         TabIndex        =   0
         Top             =   375
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   117768193
         CurrentDate     =   37518
      End
      Begin VB.Label Label2 
         Caption         =   "Tipo Doc."
         Height          =   315
         Left            =   405
         TabIndex        =   10
         Top             =   1815
         Width           =   825
      End
      Begin VB.Label Label1 
         Caption         =   "Cliente"
         Height          =   330
         Left            =   435
         TabIndex        =   9
         Top             =   1350
         Width           =   945
      End
      Begin VB.Label lbl 
         Caption         =   "Desde"
         Height          =   255
         Index           =   1
         Left            =   450
         TabIndex        =   8
         Top             =   450
         Width           =   855
      End
      Begin VB.Label lbl 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   2
         Left            =   450
         TabIndex        =   7
         Top             =   900
         Width           =   825
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   360
      Index           =   1
      Left            =   3345
      TabIndex        =   5
      Top             =   3000
      Width           =   1170
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   360
      Index           =   0
      Left            =   1710
      TabIndex        =   4
      Top             =   3000
      Width           =   1170
   End
   Begin Crystal.CrystalReport oCrystalReport 
      Left            =   0
      Top             =   2850
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "frmRepListadoDocumentos"
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
   Ctr_Cliente.conexion VGCNx
   Ctr_TipoDoc.conexion VGCNx
End Sub

Private Sub cmdCancelar_Click(Index As Integer)
   Unload Me
End Sub

Private Sub cmdAceptar_Click(Index As Integer)
   Call Imprimir
End Sub

Sub Imprimir()
'FIXIT: Declare 'arrform' con un tipo de datos de enlace en tiempo de compilación          FixIT90210ae-R1672-R1B8ZE
Dim arrform(3) As Variant, arrparm(5) As Variant
Dim NombreRep As String, CadOrden As String
Dim NombreSubRep As String
Dim NombrePC As String
Dim ValorRango As String
Dim i As Integer
   Randomize   ' Inicializa el generador de números aleatorios.
'FIXIT: Reemplazar la función 'Str' con la función 'Str$'.                                 FixIT90210ae-R9757-R1B8ZE
   NombrePC = RTrim$(Str(CLng(Rnd * 10000000)))
   arrform(0) = "Empresa='" & Right$(RTrim$(g_DetalleEmpresa), Len(RTrim$(g_DetalleEmpresa)) - 3) & "'"
   arrform(1) = "Desde='" & Format(DTDesde.Value, "dd/mm/yyyy") & "'"
   arrform(2) = "Hasta='" & Format(DTHasta.Value, "dd/mm/yyyy") & "'"
   arrparm(0) = VGCNx.DefaultDatabase
   arrparm(1) = Format(DTDesde.Value, "dd/mm/yyyy")
   arrparm(2) = Format(DTHasta.Value, "dd/mm/yyyy")
   arrparm(3) = IIf(Ctr_Cliente.xclave = Empty, "%", RTrim$(Ctr_Cliente.xclave))
   arrparm(4) = IIf(Ctr_TipoDoc.xclave = Empty, "%", RTrim$(Ctr_TipoDoc.xclave))
   NombreRep = "cc_DocVarios.rpt"
   NombreSubRep = "cc_DocVarios_sub.rpt"
   CadOrden = ""
   Call ImpresionRptProc(NombreRep, arrform, arrparm, CadOrden, "Documentos ")
End Sub


Private Sub DTDesde_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub DTHasta_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys "{TAB}"
End Sub


