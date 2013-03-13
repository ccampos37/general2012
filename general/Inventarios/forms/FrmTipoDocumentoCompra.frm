VERSION 5.00
Object = "{FED6C0D4-BBAF-48FE-B6CE-FFC87978CBAE}#215.0#0"; "Controles.ocx"
Begin VB.Form FrmTipoDocumentoCompra 
   Caption         =   "Form2"
   ClientHeight    =   9495
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9855
   LinkTopic       =   "Form2"
   ScaleHeight     =   9495
   ScaleWidth      =   9855
   StartUpPosition =   3  'Windows Default
   Begin UMantenimiento.TablasBasicas TablasBasicas1 
      Height          =   9135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   16113
   End
End
Attribute VB_Name = "FrmTipoDocumentoCompra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim a_Array(0 To 12, 0 To 12)

Private Sub Form_Load()
   'Nombre Campos:
   a_Array(0, 0) = "tipoordencodigo"
   a_Array(0, 1) = "tipoordendescripcion"
   a_Array(0, 2) = "tipoordennumeracion"
   a_Array(0, 3) = "flagrequerimientos"
   a_Array(0, 4) = "ordendebienes"
   a_Array(0, 5) = "usuariocodigo"
   a_Array(0, 6) = "fechaact"
   
   'Etiquetas:
   a_Array(1, 0) = "Codigo tipo de orden"
   a_Array(1, 1) = "descripcion tipo de orden"
   a_Array(1, 2) = "nnumeracion"
   a_Array(1, 3) = "flag de requerimientos"
   a_Array(1, 4) = "flag de orden de bienes"
   a_Array(1, 5) = Empty
   a_Array(1, 6) = Empty
   
   'Tipo de Dato:
   
   a_Array(2, 0) = "C"
   a_Array(2, 1) = "C"
   a_Array(2, 2) = "N"
   a_Array(2, 3) = "B"
   a_Array(2, 4) = "B"
   a_Array(2, 5) = "C"
   a_Array(2, 6) = "D"
   
   'Ancho de campo:
   
   a_Array(3, 0) = 10
   a_Array(3, 1) = 30
   a_Array(3, 2) = 6
   a_Array(3, 3) = 1
   a_Array(3, 4) = 1
   a_Array(3, 5) = 8
   a_Array(3, 6) = 10
   
   'Campo Clave:
   
   a_Array(4, 0) = True
   a_Array(4, 1) = False
   a_Array(4, 2) = False
   a_Array(4, 3) = False
   a_Array(4, 4) = False
   a_Array(4, 5) = False
   a_Array(4, 6) = False
   
   'Valores Ingresados por el Sistema:
   
   a_Array(5, 0) = Empty
   a_Array(5, 1) = Empty
   a_Array(5, 2) = Empty
   a_Array(5, 3) = Empty
   a_Array(5, 4) = Empty
   a_Array(5, 5) = VGUsuario
   a_Array(5, 6) = Date

   'Permite Nulos:
   
   a_Array(6, 0) = False
   a_Array(6, 1) = False
   a_Array(6, 2) = False
   a_Array(6, 3) = False
   a_Array(6, 4) = False
   a_Array(6, 5) = False
   a_Array(6, 6) = False
   
   TablasBasicas1.Conexion = VGCNx
   TablasBasicas1.NombreTabla = "co_tipodeorden"
'   TablasBasicas1.TituloForm = "Tipos de Documentos de Ordenes"
'   TablasBasicas1.Filtro = ""
   TablasBasicas1.Arreglo = a_Array
   TablasBasicas1.Setear_Controles
   TablasBasicas1.Obtener_Campos
   TablasBasicas1.cargar_datos
   
End Sub

Private Sub TablasBasicas1_Click(indice As Variant)
  
If indice = 3 Then Call Impresion("al_TipodeOrdenes.rpt")


End Sub

