VERSION 5.00
Object = "{272034D2-AC5F-11D6-810B-0050BAA91DB7}#18.0#0"; "mTablaBasica.ocx"
Begin VB.Form FrmMntTipodeOrden 
   Caption         =   "Form1"
   ClientHeight    =   6375
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8370
   LinkTopic       =   "Form1"
   ScaleHeight     =   6375
   ScaleWidth      =   8370
   StartUpPosition =   3  'Windows Default
   Begin MantTablaBasica.mTablaBasica mTablaBasica1 
      Height          =   6090
      Left            =   135
      TabIndex        =   0
      Top             =   135
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   10742
   End
End
Attribute VB_Name = "FrmMntTipodeOrden"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim a_Array(0 To 12, 0 To 12)

Private Sub Form_Load()
  'CentrarForm MDIPrincipal, Me
      
   'Nombre Campos:
   a_Array(0, 0) = "tipoordencodigo"
   a_Array(0, 1) = "tipoordendescripcion"
   a_Array(0, 2) = "ordendebienes"
   a_Array(0, 3) = "flagrequerimientosordenes"
   a_Array(0, 5) = "AprobacionGerencia"
   a_Array(0, 4) = "tipoordennumeracion"
   a_Array(0, 6) = "usuariocodigo"
   a_Array(0, 7) = "fechaact"
   
   'Etiquetas:
   a_Array(1, 0) = "Código"
   a_Array(1, 1) = "Descripcion"
   a_Array(1, 2) = "(B)ien/(S)erv."
   a_Array(1, 3) = "Para Req."
   a_Array(1, 5) = "Aprob.Geren."
   a_Array(1, 4) = "Nro.Correlativo"
   a_Array(1, 6) = Empty
   a_Array(1, 7) = Empty
   
   'Tipo de Dato:
   a_Array(2, 0) = "C"
   a_Array(2, 1) = "C"
   a_Array(2, 2) = "C"
   a_Array(2, 3) = "C"
   a_Array(2, 5) = "B"
   a_Array(2, 4) = "N"
   a_Array(2, 6) = "C"
   a_Array(2, 7) = "D"
   
   'Ancho de campo:
   a_Array(3, 0) = 10
   a_Array(3, 1) = 30
   a_Array(3, 3) = 1
   a_Array(3, 4) = 6

   
   'Campo Clave:
   a_Array(4, 0) = True
   a_Array(4, 1) = False
   a_Array(4, 2) = False
   a_Array(4, 3) = False
   a_Array(4, 4) = False
   
   'Valores Ingresados por el Sistema:
   
   a_Array(5, 0) = Empty
   a_Array(5, 1) = Empty
   a_Array(5, 2) = Empty
   a_Array(5, 3) = Empty
   a_Array(5, 4) = Empty
   a_Array(5, 5) = Empty
   a_Array(5, 6) = VGUsuario
   a_Array(5, 7) = Date

   
   'a_Array(5, 4) = Format(Now, "aaaa-mm-dd hh:mm:ss.000")
   'Permite Nulos:
   
   a_Array(6, 0) = False
   a_Array(6, 1) = False
   a_Array(6, 2) = False
   a_Array(6, 3) = False
   a_Array(6, 4) = False
   a_Array(6, 5) = False
   a_Array(6, 6) = False
   a_Array(6, 7) = False
   
   mTablaBasica1.conexion = VGCNx
   mTablaBasica1.NombreTabla = "co_tipodeorden"
   mTablaBasica1.TituloForm = "Tipo de Ordenes"
   mTablaBasica1.filtro = ""
   mTablaBasica1.Arreglo = a_Array
   mTablaBasica1.Setear_Controles
   mTablaBasica1.Obtener_Campos
   mTablaBasica1.cargar_datos
End Sub
Private Sub mTablaBasica1_Click(indice As Variant)
  If indice = 3 Then Call impresion("co_tipodeorden.rpt")
End Sub

