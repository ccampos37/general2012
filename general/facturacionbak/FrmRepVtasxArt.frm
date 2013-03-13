VERSION 5.00
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form FrmRepVtasxArt 
   Caption         =   "Ventas por Articulo"
   ClientHeight    =   3015
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5925
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   5925
   StartUpPosition =   3  'Windows Default
   Begin Crystal.CrystalReport oCrystalReport 
      Left            =   240
      Top             =   2040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComCtl2.DTPicker DTHasta 
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   1680
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
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
      Format          =   61603841
      CurrentDate     =   37518
   End
   Begin MSComCtl2.DTPicker DTDesde 
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   1200
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
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
      Format          =   61603841
      CurrentDate     =   37518
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   3000
      TabIndex        =   3
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   1320
      TabIndex        =   2
      Top             =   2400
      Width           =   1215
   End
   Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayualmacen 
      Height          =   375
      Left            =   1320
      TabIndex        =   8
      Top             =   240
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   661
      XcodMaxLongitud =   0
      xcodwith        =   100
      NomTabla        =   "tabalm"
      TituloAyuda     =   "Almacenes"
      ListaCampos     =   "TAALMA(1),TADESCRI(1)"
      XcodCampo       =   "TAALMA"
      XListCampo      =   "TADESCRI"
      ListaCamposDescrip=   "Codigo,Descripcion"
      ListaCamposText =   "TAALMA,TADESCRI"
      Requerido       =   0   'False
   End
   Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyuProducto 
      Height          =   375
      Left            =   1320
      TabIndex        =   9
      Top             =   720
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   661
      XcodMaxLongitud =   11
      xcodwith        =   800
      NomTabla        =   "maeart"
      TituloAyuda     =   "Busqueda de Codigos de producto"
      ListaCampos     =   "acodigo(1),adescri(1)"
      XcodCampo       =   "acodigo"
      XListCampo      =   "adescri"
      ListaCamposDescrip=   "Codigo,Descripcion"
      ListaCamposText =   "acodigo,adescri"
      Requerido       =   0   'False
   End
   Begin VB.Label lbl 
      Caption         =   "Articulo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   7
      Top             =   720
      Width           =   855
   End
   Begin VB.Label lbl 
      Caption         =   "Almacen"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label lbl 
      Caption         =   "Desde"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   5
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label lbl 
      Caption         =   "Hasta"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   4
      Top             =   1680
      Width           =   615
   End
End
Attribute VB_Name = "FrmRepVtasxArt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adll As New dllgeneral.dll_general
''''''''''''''''''''''''''''''''''''''''''''''''''''
'Agregar:
Dim busca As New dll_apisgen.dll_apis
''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub cmdAceptar_Click(Index As Integer)
Dim aform(3) As Variant
Dim aparam(6) As Variant

 If DTDesde > DTHasta Then
     MsgBox "Fecha Desde debe ser mayor a Fecha Hasta", vbInformation, "AVISO"
     Exit Sub
  End If
aform(0) = "Desde='" & DTDesde & "'"
aform(1) = "Hasta='" & DTHasta & "'"
If Ctr_Ayualmacen.xclave <> "" Then
   aform(2) = "Almacen='" & Ctr_Ayualmacen.xclave & "'"
 Else
   aform(2) = "Almacen='TODOS'"
End If
If Ctr_AyuProducto.xclave <> "" Then
   aform(3) = "Articulo='" & Ctr_AyuProducto.xclave & "'"
   aparam(3) = Ctr_AyuProducto.xclave
 Else
   aform(3) = "Articulo='TODOS'"
   aparam(3) = "%%"
End If
aparam(0) = VGCNx.DefaultDatabase
aparam(1) = VGParametros.empresacodigo
aparam(2) = IIf(Ctr_Ayualmacen.xclave = "", "%%", Ctr_Ayualmacen.xclave)
aparam(4) = DTDesde
aparam(5) = DTHasta

Call ImpresionRptProc("vt_VtasxArticulo.rpt", aform, aparam, , "Ventas por Articulo")

End Sub

Private Sub cmdCancelar_Click(Index As Integer)
    Unload Me
End Sub

Private Sub DtDesde_KeyDown(KeyCode As Integer, Shift As Integer)
      If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub DtHasta_KeyDown(KeyCode As Integer, Shift As Integer)
      If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
    MostrarForm Me, "C2"
    Call Ctr_Ayualmacen.Conexion(VGCNx)
    Call Ctr_AyuProducto.Conexion(VGCNx)
'      Ctr_Ayuda3.Filtro = " empresacodigo='" & VGParametros.empresacodigo & "' and puntovtacodigo='" & VGParametros.puntovta & "'"
    DTDesde = Date
    DTHasta = Date
End Sub

Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
