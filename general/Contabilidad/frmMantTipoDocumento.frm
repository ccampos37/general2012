VERSION 5.00
Object = "{272034D2-AC5F-11D6-810B-0050BAA91DB7}#18.0#0"; "mTablaBasica.ocx"
Begin VB.Form frmMantTipoDocumento 
   Caption         =   "Tipo de Documentos"
   ClientHeight    =   6105
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7560
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6105
   ScaleWidth      =   7560
   Begin MantTablaBasica.mTablaBasica mTablaBasica1 
      Height          =   6150
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7545
      _ExtentX        =   13309
      _ExtentY        =   10848
   End
End
Attribute VB_Name = "frmMantTipoDocumento"
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
   a_Array(0, 0) = "documentocodigo"
   a_Array(0, 1) = "documentodescripcion"
   a_Array(0, 2) = "documentoref"
   a_Array(0, 3) = "documentoregcompras"
   a_Array(0, 4) = "documentoregventas"
   a_Array(0, 5) = "documentoregletrasxcobrar"
   a_Array(0, 6) = "documentoregletrasxpagar"
   a_Array(0, 7) = "documentonotacredito"
   
   a_Array(0, 8) = "usuariocodigo"
   a_Array(0, 9) = "fechaact"
   'Etiquetas:
   a_Array(1, 0) = "Código"
   a_Array(1, 1) = "Descripción"
   a_Array(1, 2) = "Doc. Referencia"
   a_Array(1, 3) = "Reg. Compras"
   a_Array(1, 4) = "Reg. Ventas"
   a_Array(1, 5) = "Letras x Cobrar"
   a_Array(1, 6) = "Letras x Pagar"
   a_Array(1, 7) = "Nota Crédito"
   
   a_Array(1, 8) = Empty
   a_Array(1, 9) = Empty
   'Tipo de Dato:
   a_Array(2, 0) = "C"
   a_Array(2, 1) = "C"
   a_Array(2, 2) = "B"
   
   a_Array(2, 3) = "B"
   a_Array(2, 4) = "B"
   a_Array(2, 5) = "B"
   a_Array(2, 6) = "B"
   a_Array(2, 7) = "B"
   
   a_Array(2, 8) = "C"
   a_Array(2, 9) = "D"
   'Ancho de campo:
   a_Array(3, 0) = 2
   a_Array(3, 1) = 40
   a_Array(3, 2) = 1
   
   a_Array(3, 3) = 1
   a_Array(3, 4) = 1
   a_Array(3, 5) = 1
   a_Array(3, 6) = 1
   a_Array(3, 7) = 1
   
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
   a_Array(6, 4) = False
   a_Array(6, 5) = False
   a_Array(6, 6) = False
   a_Array(6, 7) = False
   a_Array(6, 8) = False
   a_Array(6, 9) = False
   
   mTablaBasica1.conexion = VGCNx
   mTablaBasica1.NombreTabla = "gr_documento"
   mTablaBasica1.TituloForm = "Tipo de Documento"
   mTablaBasica1.Filtro = "documentocodigo<>'00'"
   mTablaBasica1.Arreglo = a_Array
   mTablaBasica1.Setear_Controles
   mTablaBasica1.Obtener_Campos
   mTablaBasica1.cargar_datos
   
End Sub

'FIXIT: Declare 'indice' con un tipo de datos de enlace en tiempo de compilación           FixIT90210ae-R1672-R1B8ZE
Private Sub mTablaBasica1_Click(indice As Variant)
  If indice = 3 Then Call Impresion("rptTipoDocumento.rpt")
End Sub

