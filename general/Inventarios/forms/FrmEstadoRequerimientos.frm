VERSION 5.00
Object = "{FED6C0D4-BBAF-48FE-B6CE-FFC87978CBAE}#215.0#0"; "Controles.ocx"
Begin VB.Form FrmEstadoRequerimientos 
   Caption         =   "Form1"
   ClientHeight    =   9240
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9435
   LinkTopic       =   "Form1"
   ScaleHeight     =   9240
   ScaleWidth      =   9435
   StartUpPosition =   3  'Windows Default
   Begin UMantenimiento.TablasBasicas TablasBasicas1 
      Height          =   8895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   15690
   End
End
Attribute VB_Name = "FrmEstadoRequerimientos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim a_Array(0 To 12, 0 To 12)
Private Sub Form_Load()
   central Me
     
   a_Array(0, 0) = "estadooccodigo"
   a_Array(0, 1) = "estadoocdescripcion"
   a_Array(0, 2) = "estadoocatendido"
   
   'Etiquetas:
   a_Array(1, 0) = "Código"
   a_Array(1, 1) = "Decripcion"
   a_Array(1, 2) = "Estado de Atendido"
   
   'Tipo de Dato:
   
   a_Array(2, 0) = "C"
   a_Array(2, 1) = "C"
   a_Array(2, 2) = "B"
   
   'Ancho de campo:
   a_Array(3, 0) = 1
   a_Array(3, 1) = 40
   a_Array(3, 2) = 1
   
   'Campo Clave:
   
   a_Array(4, 0) = True
   a_Array(4, 1) = False
   a_Array(4, 2) = False
   
   'Valores Ingresados por el Sistema:
   
   a_Array(5, 0) = Empty
   a_Array(5, 1) = Empty
   a_Array(5, 2) = 0
   
   'a_Array(5, 4) = Format(Now, "aaaa-mm-dd hh:mm:ss.000")
   'Permite Nulos:
   
   a_Array(6, 0) = False
   a_Array(6, 1) = False
   a_Array(6, 2) = False
   
   TablasBasicas1.Conexion = VGCNx
   TablasBasicas1.NombreTabla = "co_estadoRequerimiento"
   TablasBasicas1.Arreglo = a_Array
   TablasBasicas1.Setear_Controles
   TablasBasicas1.Obtener_Campos
   TablasBasicas1.cargar_datos
 
End Sub

Private Sub mTablaBasica1_Click(indice As Variant)
  If indice = 3 Then Call Impresion("co_EstadodeRequerimientos.rpt")
End Sub




