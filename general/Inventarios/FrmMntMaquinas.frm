VERSION 5.00
Object = "{272034D2-AC5F-11D6-810B-0050BAA91DB7}#18.0#0"; "mtablabasica.ocx"
Begin VB.Form FrmMntMaquinas 
   Caption         =   "Form1"
   ClientHeight    =   6600
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7890
   LinkTopic       =   "Form1"
   ScaleHeight     =   6600
   ScaleWidth      =   7890
   StartUpPosition =   3  'Windows Default
   Begin MantTablaBasica.mTablaBasica mTablaBasica1 
      Height          =   6375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   11245
   End
End
Attribute VB_Name = "FrmMntMaquinas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim a_Array(0 To 12, 0 To 12)
Private Sub Form_Load()
   Me.Width = 7860: Me.Height = 6795
   mTablaBasica1.Width = 7590
   central Me
     
   'Nombre Campos:
   a_Array(0, 0) = "codigomaquina"
   a_Array(0, 1) = "descripcionmaquina"
   a_Array(0, 2) = "factormaquina"
   'Etiquetas:
   a_Array(1, 0) = "Código"
   a_Array(1, 1) = "Nombre"
   a_Array(1, 2) = "Densidad"
   'Tipo de Dato:
   a_Array(2, 0) = "C"
   a_Array(2, 1) = "C"
   a_Array(2, 2) = "N"
   'Ancho de campo:
   a_Array(3, 0) = 4
   a_Array(3, 1) = 30
   a_Array(3, 2) = 10
   'Campo Clave:
   a_Array(4, 0) = True
   a_Array(4, 1) = False
   a_Array(4, 2) = False
   'Valores Ingresados por el Sistema:
   a_Array(5, 0) = Empty
   a_Array(5, 1) = Empty
   a_Array(5, 2) = Empty
   'a_Array(5, 4) = Format(Now, "aaaa-mm-dd hh:mm:ss.000")
   'Permite Nulos:
   a_Array(6, 0) = False
   a_Array(6, 1) = False
   a_Array(6, 2) = False
   mTablaBasica1.Conexion = VGcnx
   mTablaBasica1.nombretabla = "al_tipomaquina"
   mTablaBasica1.TituloForm = "Tipo de Maquinas"
   mTablaBasica1.Arreglo = a_Array
   mTablaBasica1.Setear_Controles
   mTablaBasica1.Obtener_Campos
   mTablaBasica1.cargar_datos
 
End Sub

Private Sub TablaBasica1_Click(indice As Variant)
  If indice = 3 Then Call Impresion("al_rpttipomaquinas.rpt ")
End Sub




