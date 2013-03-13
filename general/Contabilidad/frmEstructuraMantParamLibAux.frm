VERSION 5.00
Object = "{272034D2-AC5F-11D6-810B-0050BAA91DB7}#18.0#0"; "mTablaBasica.ocx"
Begin VB.Form frmEstructuraMantParamLibAux 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Parámetros Libros Auxiliares"
   ClientHeight    =   5940
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7470
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5940
   ScaleWidth      =   7470
   Begin MantTablaBasica.mTablaBasica mTablaBasica1 
      Height          =   6135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   10821
   End
End
Attribute VB_Name = "frmEstructuraMantParamLibAux"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'FIXIT: Declare 'a_Array' con un tipo de datos de enlace en tiempo de compilación          FixIT90210ae-R1672-R1B8ZE
Dim a_Array(0 To 12, 0 To 12) As Variant

Private Sub Form_Load()
   Me.Width = 7590: Me.Height = 6390
   mTablaBasica1.Width = 7545
   'CentrarForm MDIPrincipal, Me
      
   'Nombre Campos:
   a_Array(0, 0) = "paramlibauxtipo"
   a_Array(0, 1) = "paramlibauxdescripcion"
   a_Array(0, 2) = "paramlibauxasiento"
   a_Array(0, 3) = "paramlibauxcuenta"
   a_Array(0, 4) = "paramlibauxigv"
   a_Array(0, 5) = "paramlibauxirenta"
   a_Array(0, 6) = "paramlibauxies"
   a_Array(0, 7) = "paramlibauxtiporeg"
   a_Array(0, 8) = "usuariocodigo"
   a_Array(0, 9) = "fechaact"
   'Etiquetas:
   a_Array(1, 0) = "Tipo de Reporte"
   a_Array(1, 1) = "Descripción"
   a_Array(1, 2) = "Asientos"
   a_Array(1, 3) = "Cuentas"
   a_Array(1, 4) = "IGV"
   a_Array(1, 5) = "I.RENTA / OTROS INGRESOS"
   a_Array(1, 6) = "IES / FLETE"
   a_Array(1, 7) = "Tipo de Registro"
   a_Array(1, 8) = Empty
   a_Array(1, 9) = Empty
   'Tipo de Dato:
   a_Array(2, 0) = "C"
   a_Array(2, 1) = "C"
   a_Array(2, 2) = "C"
   a_Array(2, 3) = "C"
   a_Array(2, 4) = "C"
   a_Array(2, 5) = "C"
   a_Array(2, 6) = "C"
   a_Array(2, 7) = "C"
   a_Array(2, 8) = "C"
   a_Array(2, 9) = "D"
   'Ancho de campo:
   a_Array(3, 0) = 6
   a_Array(3, 1) = 20
   a_Array(3, 2) = 40
   a_Array(3, 3) = 70
   a_Array(3, 4) = 40
   a_Array(3, 5) = 40
   a_Array(3, 6) = 40
   a_Array(3, 7) = 4
   a_Array(3, 8) = 8
   a_Array(3, 9) = Empty
   'Campo Clave:
   a_Array(4, 0) = True
   a_Array(4, 1) = False
   a_Array(4, 2) = False
   a_Array(4, 3) = False
   a_Array(4, 4) = False
   a_Array(4, 5) = False
   a_Array(4, 6) = False
   a_Array(4, 7) = False
   a_Array(4, 8) = False
   a_Array(4, 9) = False
   'Valores Ingresados por el Sistema:
   a_Array(5, 0) = Empty
   a_Array(5, 1) = Empty
   a_Array(5, 2) = Empty
   a_Array(5, 3) = Empty
   a_Array(5, 4) = Empty
   a_Array(5, 5) = Empty
   a_Array(5, 6) = Empty
   a_Array(5, 7) = Empty
   a_Array(5, 8) = VGusuario
   a_Array(5, 9) = Date
   'a_Array(5, 4) = Format(Now, "aaaa-mm-dd hh:mm:ss.000")
   'Permite Nulos:
   a_Array(6, 0) = False
   a_Array(6, 1) = False
   a_Array(6, 2) = False
   a_Array(6, 3) = False
   a_Array(6, 4) = True
   a_Array(6, 5) = True
   a_Array(6, 6) = True
   a_Array(6, 7) = True
   a_Array(6, 8) = False
   a_Array(6, 9) = False
   
   mTablaBasica1.conexion = VGCNx
   mTablaBasica1.NombreTabla = "ct_paramlibaux"
   mTablaBasica1.TituloForm = "Parámetros Libros Auxiliares"
   mTablaBasica1.Arreglo = a_Array
   mTablaBasica1.Setear_Controles
   mTablaBasica1.Obtener_Campos
   mTablaBasica1.cargar_datos
   
End Sub

'FIXIT: Declare 'indice' con un tipo de datos de enlace en tiempo de compilación           FixIT90210ae-R1672-R1B8ZE
Private Sub mTablaBasica1_Click(indice As Variant)
  If indice = 3 Then Call Impresion("rptEstruMantParLibAux.rpt")
End Sub
