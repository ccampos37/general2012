VERSION 5.00
Object = "{272034D2-AC5F-11D6-810B-0050BAA91DB7}#18.0#0"; "mTablaBasica.ocx"
Begin VB.Form frmEstructuraMantParametrosGastos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Parámetros Gastos"
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
Attribute VB_Name = "frmEstructuraMantParametrosGastos"
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
   a_Array(0, 0) = "paramgastoslinutil"
   a_Array(0, 1) = "paramgastosctautil"
   a_Array(0, 2) = "paramgastoslinventa"
   a_Array(0, 3) = "paramgastoslinadmin"
   'a_Array(0, 4) = "paramgastoslindiv"
   a_Array(0, 4) = "paramgastosactivo"
   a_Array(0, 5) = "paramgastosgastoadmin"
   a_Array(0, 6) = "paramgastosgastoventa"
   a_Array(0, 7) = "paramgastosgastoprod"
   a_Array(0, 8) = "paramgastosgastofinan"
   a_Array(0, 9) = "paramgastosgastodiv"
   a_Array(0, 10) = "usuariocodigo"
   a_Array(0, 11) = "fechaact"
   'Etiquetas:
   a_Array(1, 0) = "Línea de Utilidad"
   a_Array(1, 1) = "Cuenta de Utilidad"
   a_Array(1, 2) = "Linea de Ventas"
   a_Array(1, 3) = "Linea de Administración"
   a_Array(1, 4) = "Activo"
   a_Array(1, 5) = "Gastos de Adm."
   a_Array(1, 6) = "Gastos de Ventas"
   a_Array(1, 7) = "Gastos de Producción"
   a_Array(1, 8) = "Gastos Financieros"
   a_Array(1, 9) = "Gastos Diversos"
   a_Array(1, 10) = Empty
   a_Array(1, 11) = Empty
   'Tipo de Dato:
   a_Array(2, 0) = "N"
   a_Array(2, 1) = "N"
   a_Array(2, 2) = "N"
   a_Array(2, 3) = "N"
   a_Array(2, 4) = "N"
   a_Array(2, 5) = "N"
   a_Array(2, 6) = "N"
   a_Array(2, 7) = "N"
   a_Array(2, 8) = "N"
   a_Array(2, 9) = "N"
   a_Array(2, 10) = "C"
   a_Array(2, 11) = "D"
   'Ancho de campo:
   a_Array(3, 0) = 6
   a_Array(3, 1) = 2
   a_Array(3, 2) = 2
   a_Array(3, 3) = 2
   a_Array(3, 4) = 2
   a_Array(3, 5) = 2
   a_Array(3, 6) = 2
   a_Array(3, 7) = 2
   a_Array(3, 8) = 2
   a_Array(3, 9) = 2
   a_Array(3, 10) = 8
   a_Array(3, 11) = Empty
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
   a_Array(4, 10) = False
   a_Array(4, 11) = False
   'Valores Ingresados por el Sistema:
   a_Array(5, 0) = Empty
   a_Array(5, 1) = Empty
   a_Array(5, 2) = Empty
   a_Array(5, 3) = Empty
   a_Array(5, 4) = Empty
   a_Array(5, 5) = Empty
   a_Array(5, 6) = Empty
   a_Array(5, 7) = Empty
   a_Array(5, 8) = Empty
   a_Array(5, 9) = Empty
   a_Array(5, 10) = VGusuario
   a_Array(5, 11) = Date
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
   a_Array(6, 8) = False
   a_Array(6, 9) = False
   a_Array(6, 10) = False
   a_Array(6, 11) = False
   
   mTablaBasica1.conexion = VGCNx
   mTablaBasica1.NombreTabla = "ct_paramgastos"
   mTablaBasica1.TituloForm = "Parámetros de Gastos"
   mTablaBasica1.Arreglo = a_Array
   mTablaBasica1.Setear_Controles
   mTablaBasica1.Obtener_Campos
   mTablaBasica1.cargar_datos
   
End Sub

'FIXIT: Declare 'indice' con un tipo de datos de enlace en tiempo de compilación           FixIT90210ae-R1672-R1B8ZE
Private Sub mTablaBasica1_Click(indice As Variant)
  If indice = 3 Then Call Impresion("rptEstruMantParGastos.rpt")
End Sub
