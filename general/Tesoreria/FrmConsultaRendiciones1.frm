VERSION 5.00
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Begin VB.Form FrmConsultarendiciones 
   Caption         =   "Form1"
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8985
   LinkTopic       =   "Form1"
   ScaleHeight     =   6600
   ScaleWidth      =   8985
   StartUpPosition =   3  'Windows Default
   Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyudaOficina 
      Height          =   300
      Left            =   990
      TabIndex        =   0
      Top             =   0
      Width           =   4860
      _ExtentX        =   8573
      _ExtentY        =   529
      XcodMaxLongitud =   3
      xcodwith        =   400
      NomTabla        =   "cp_oficina"
      TituloAyuda     =   "Ayuda de Caja"
      ListaCampos     =   "vendedorcodigo(1),vendedornombres(1)"
      XcodCampo       =   "vendedorcodigo"
      XListCampo      =   "vendedornombres"
      ListaCamposDescrip=   "Código,Descripción"
      ListaCamposText =   "vendedorcodigo,vendedornombres"
   End
   Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyudaMoneda 
      Height          =   315
      Left            =   1050
      TabIndex        =   1
      Top             =   1050
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   556
      XcodMaxLongitud =   2
      xcodwith        =   300
      NomTabla        =   "gr_moneda"
      TituloAyuda     =   "Busqueda de Moneda"
      ListaCampos     =   "monedacodigo(1),monedadescripcion(1)"
      XcodCampo       =   "monedacodigo"
      XListCampo      =   "monedadescripcion"
      ListaCamposDescrip=   "Código,Descripción"
      ListaCamposText =   "monedacodigo,monedadescripcion"
   End
   Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyudaCaja 
      Height          =   315
      Left            =   975
      TabIndex        =   2
      Top             =   450
      Width           =   4860
      _ExtentX        =   8573
      _ExtentY        =   556
      XcodMaxLongitud =   11
      xcodwith        =   400
      NomTabla        =   "te_codigocaja"
      TituloAyuda     =   "Busqueda de Caja"
      ListaCampos     =   "cajacodigo(1),cajadescripcion(1)"
      XcodCampo       =   "cajacodigo"
      XListCampo      =   "cajadescripcion"
      ListaCamposDescrip=   "Codigo,Descripcion"
      ListaCamposText =   "cajacodigo,cajadescripcion"
   End
   Begin VB.Label Label4 
      Caption         =   "Oficina"
      Height          =   255
      Index           =   0
      Left            =   30
      TabIndex        =   5
      Top             =   165
      Width           =   885
   End
   Begin VB.Label lbMon 
      Caption         =   "Moneda : "
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   1095
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "Cod. Caja"
      Height          =   255
      Index           =   1
      Left            =   15
      TabIndex        =   3
      Top             =   450
      Width           =   885
   End
End
Attribute VB_Name = "FrmConsultarendiciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Ctr_AyudaCaja_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
Dim rsql As New ADODB.Recordset
SQL = " select * from te_codigocaja where cajacodigo='" & Ctr_AyudaCaja.xclave & "'"
    Set rsql = VGcnx.Execute(SQL)
  '  Call Listar

End Sub

Private Sub Form_Load()
    Call Ctr_AyudaOficina.Conexion(VGcnx)
    Call Ctr_AyudaCaja.Conexion(VGcnx): Ctr_AyudaCaja.Filtro = " cajarendiciones=1 "
    Call Ctr_AyudaMoneda.Conexion(VGcnx)
End Sub
