VERSION 5.00
Object = "{272034D2-AC5F-11D6-810B-0050BAA91DB7}#18.0#0"; "mTablaBasica.ocx"
Begin VB.Form FrmTipocuenta 
   Caption         =   "Form1"
   ClientHeight    =   6765
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   8235
   LinkTopic       =   "Form1"
   ScaleHeight     =   6765
   ScaleWidth      =   8235
   StartUpPosition =   3  'Windows Default
   Begin MantTablaBasica.mTablaBasica mTablaBasica1 
      Height          =   6375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   11245
   End
End
Attribute VB_Name = "FrmTipocuenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'FIXIT: Declare 'a_Array as variant' con un tipo de datos de enlace en tiempo de compilación     FixIT90210ae-R1672-R1B8ZE
Dim a_Array(0 To 12, 0 To 12) As Variant

Private Sub Form_Load()
   CentrarForm MDIPrincipal, Me
      
   'Nombre Campos:
   a_Array(0, 0) = "tipocuentacodigo"
   a_Array(0, 1) = "tipocuentadescripcion"
   a_Array(0, 2) = "tipocuentaabreviatura"
   'Etiquetas:
   a_Array(1, 0) = "Código"
   a_Array(1, 1) = "Descripción"
   a_Array(1, 2) = "Abreviatura"
   'Tipo de Dato:
   a_Array(2, 0) = "C"
   a_Array(2, 1) = "C"
   a_Array(2, 2) = "C"
   'Ancho de campo:
   a_Array(3, 0) = 2
   a_Array(3, 1) = 25
   a_Array(3, 2) = 3
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
   
   mTablaBasica1.conexion = VGCNx
   mTablaBasica1.NombreTabla = "ct_tipocuenta"
   mTablaBasica1.TituloForm = "Tipos de Cuentas"
   mTablaBasica1.Arreglo = a_Array
   mTablaBasica1.Setear_Controles
   mTablaBasica1.Obtener_Campos
   mTablaBasica1.cargar_datos
   
End Sub

'FIXIT: Declare 'indice' con un tipo de datos de enlace en tiempo de compilación           FixIT90210ae-R1672-R1B8ZE
Private Sub mTablaBasica1_Click(indice As Variant)
    If indice = 3 Then
       Call Impresion("ct_tipocuenta.rpt")
    End If
End Sub


