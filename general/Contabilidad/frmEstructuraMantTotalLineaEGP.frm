VERSION 5.00
Object = "{272034D2-AC5F-11D6-810B-0050BAA91DB7}#18.0#0"; "mTablaBasica.ocx"
Begin VB.Form frmEstructuraMantTotalLineaEGP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Totalizador Ganacias y Pérdidas"
   ClientHeight    =   5955
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7440
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   7440
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
Attribute VB_Name = "frmEstructuraMantTotalLineaEGP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'FIXIT: Declare 'a_Array as variant' con un tipo de datos de enlace en tiempo de compilación     FixIT90210ae-R1672-R1B8ZE
Dim a_Array(0 To 12, 0 To 12) As Variant

Private Sub Form_Load()
   Me.Width = 7590: Me.Height = 6390
   mTablaBasica1.Width = 7545
   'CentrarForm MDIPrincipal, Me
      
   'Nombre Campos:
   a_Array(0, 0) = "totalizadorgplinea"
   a_Array(0, 1) = "totalizadorgpformulames"
   a_Array(0, 2) = "totalizadorgpformulaacum"
   a_Array(0, 3) = "totalizadorgpantig"
   a_Array(0, 4) = "usuariocodigo"
   a_Array(0, 5) = "fechaact"
   'Etiquetas:
   a_Array(1, 0) = "Línea"
   a_Array(1, 1) = "Formula Mensual"
   a_Array(1, 2) = "Formula Acumulado"
   a_Array(1, 3) = "Antiguedad"
   a_Array(1, 4) = Empty
   a_Array(1, 5) = Empty
   'Tipo de Dato:
   a_Array(2, 0) = "C"
   a_Array(2, 1) = "C"
   a_Array(2, 2) = "C"
   a_Array(2, 3) = "C"
   a_Array(2, 4) = "C"
   a_Array(2, 5) = "D"
   'Ancho de campo:
   a_Array(3, 0) = 6
   a_Array(3, 1) = 90
   a_Array(3, 2) = 90
   a_Array(3, 3) = 6
   a_Array(3, 4) = 8
   a_Array(3, 5) = Empty
   'Campo Clave:
   a_Array(4, 0) = True
   a_Array(4, 1) = False
   a_Array(4, 2) = False
   a_Array(4, 3) = False
   a_Array(4, 4) = False
   a_Array(4, 5) = False
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
   mTablaBasica1.NombreTabla = "ct_totalizadorgp"
   mTablaBasica1.TituloForm = "Totalizador de Ganancias y Pérdidas"
   mTablaBasica1.Arreglo = a_Array
   mTablaBasica1.Setear_Controles
   mTablaBasica1.Obtener_Campos
   mTablaBasica1.cargar_datos
   
End Sub

'FIXIT: Declare 'indice' con un tipo de datos de enlace en tiempo de compilación           FixIT90210ae-R1672-R1B8ZE
Private Sub mTablaBasica1_Click(indice As Variant)
  If indice = 3 Then Call Impresion("rptEstruMantTotalLinEGP.rpt")
End Sub
