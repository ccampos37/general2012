VERSION 5.00
Object = "{272034D2-AC5F-11D6-810B-0050BAA91DB7}#18.0#0"; "mTablaBasica.ocx"
Begin VB.Form frmEstructuraMantRatiosFinan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ratios Financieros"
   ClientHeight    =   5955
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7440
   LinkTopic       =   "Form1"
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
Attribute VB_Name = "frmEstructuraMantRatiosFinan"
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
   a_Array(0, 0) = "ratiosfinanlinea"
   a_Array(0, 1) = "ratiosfinannivel1"
   a_Array(0, 2) = "ratiosfinandescrip1"
   a_Array(0, 3) = "ratiosfinandescrip2"
   a_Array(0, 4) = "ratiosfinanformula"
   a_Array(0, 5) = "usuariocodigo"
   a_Array(0, 6) = "fechaact"
   'Etiquetas:
   a_Array(1, 0) = "Línea"
   a_Array(1, 1) = "Nivel 1"
   a_Array(1, 2) = "Descripción 1"
   a_Array(1, 3) = "Descripción 2"
   a_Array(1, 4) = "Formula"
   a_Array(1, 5) = Empty
   a_Array(1, 6) = Empty
   'Tipo de Dato:
   a_Array(2, 0) = "C"
   a_Array(2, 1) = "C"
   a_Array(2, 2) = "C"
   a_Array(2, 3) = "C"
   a_Array(2, 4) = "C"
   a_Array(2, 5) = "C"
   a_Array(2, 6) = "D"
   'Ancho de campo:
   a_Array(3, 0) = 6
   a_Array(3, 1) = 2
   a_Array(3, 2) = 110
   a_Array(3, 3) = 20
   a_Array(3, 4) = 120
   a_Array(3, 5) = 8
   a_Array(3, 6) = Empty
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
   a_Array(5, 5) = VGusuario
   a_Array(5, 6) = Date
   'a_Array(5, 4) = Format(Now, "aaaa-mm-dd hh:mm:ss.000")
   'Permite Nulos:
   a_Array(6, 0) = False
   a_Array(6, 1) = False
   a_Array(6, 2) = False
   a_Array(6, 3) = False
   a_Array(6, 4) = False
   a_Array(6, 5) = False
   a_Array(6, 6) = False
   
   mTablaBasica1.conexion = VGCNx
   mTablaBasica1.NombreTabla = "ct_ratiosfinan"
   mTablaBasica1.TituloForm = "Estructura de Ratios Financieros"
   mTablaBasica1.Arreglo = a_Array
   mTablaBasica1.Setear_Controles
   mTablaBasica1.Obtener_Campos
   mTablaBasica1.cargar_datos
   
End Sub

'FIXIT: Declare 'indice' con un tipo de datos de enlace en tiempo de compilación           FixIT90210ae-R1672-R1B8ZE
Private Sub mTablaBasica1_Click(indice As Variant)
  If indice = 3 Then Call Impresion("rptEstruMantRatFinan.rpt")
End Sub
