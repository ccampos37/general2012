VERSION 5.00
Object = "{272034D2-AC5F-11D6-810B-0050BAA91DB7}#18.0#0"; "mTablaBasica.ocx"
Begin VB.Form frmMantOperacion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Operaci�n"
   ClientHeight    =   6240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7725
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   7725
   Begin MantTablaBasica.mTablaBasica mTablaBasica1 
      Height          =   6150
      Left            =   -30
      TabIndex        =   0
      Top             =   0
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   10848
   End
End
Attribute VB_Name = "frmMantOperacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'FIXIT: Declare 'a_Array' con un tipo de datos de enlace en tiempo de compilaci�n          FixIT90210ae-R1672-R1B8ZE
Dim a_Array(0 To 12, 0 To 12) As Variant

Private Sub Form_Load()
  'CentrarForm MDIPrincipal, Me
      
   'Nombre Campos:
   a_Array(0, 0) = "operacioncodigo"
   a_Array(0, 1) = "operaciondescripcion"
   a_Array(0, 2) = "operaciondocumentoanulado"
   a_Array(0, 3) = "facturacionanticipada"
   a_Array(0, 4) = "usuariocodigo"
   a_Array(0, 5) = "fechaact"
   
   'Etiquetas:
   a_Array(1, 0) = "C�digo"
   a_Array(1, 1) = "Descripcion"
   a_Array(1, 2) = "Permite Doc.Anulado"
   a_Array(1, 3) = "Compensa fact.anticipada"
   a_Array(1, 4) = Empty
   a_Array(1, 5) = Empty
   
   'Tipo de Dato:
   a_Array(2, 0) = "C"
   a_Array(2, 1) = "C"
   a_Array(2, 2) = "B"
   a_Array(2, 3) = "B"
   a_Array(2, 4) = "C"
   a_Array(2, 5) = "D"
   
   'Ancho de campo:
   a_Array(3, 0) = 2
   a_Array(3, 1) = 25
   a_Array(3, 3) = 8
   a_Array(3, 4) = Empty

   
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
   a_Array(5, 4) = VGusuario
   a_Array(5, 5) = Date

   
   'a_Array(5, 4) = Format(Now, "aaaa-mm-dd hh:mm:ss.000")
   'Permite Nulos:
   
   a_Array(6, 0) = False
   a_Array(6, 1) = False
   a_Array(6, 2) = False
   a_Array(6, 3) = False
   a_Array(6, 4) = False
   a_Array(6, 5) = False
   
   mTablaBasica1.conexion = VGCNx
   mTablaBasica1.NombreTabla = "ct_operacion"
   mTablaBasica1.TituloForm = "Operaci�n"
   mTablaBasica1.Filtro = "operacioncodigo<>'00'"
   mTablaBasica1.Arreglo = a_Array
   mTablaBasica1.Setear_Controles
   mTablaBasica1.Obtener_Campos
   mTablaBasica1.cargar_datos
End Sub
'FIXIT: Declare 'indice' con un tipo de datos de enlace en tiempo de compilaci�n           FixIT90210ae-R1672-R1B8ZE
Private Sub mTablaBasica1_Click(indice As Variant)
  If indice = 3 Then Call Impresion("rptOperacion.rpt")
End Sub
