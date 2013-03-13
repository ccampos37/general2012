VERSION 5.00
Object = "{272034D2-AC5F-11D6-810B-0050BAA91DB7}#18.0#0"; "mTablaBasica.ocx"
Begin VB.Form frmMantTipoAnalitico 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tipo Análitico"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7515
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   7515
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
Attribute VB_Name = "frmMantTipoAnalitico"
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
   CentrarForm MDIPrincipal, Me
      
   'Nombre Campos:
   a_Array(0, 0) = "tipoanaliticocodigo"
   a_Array(0, 1) = "tipoanaliticodescripcion"
   a_Array(0, 2) = "usuariocodigo"
   a_Array(0, 3) = "fechaact"
   'Etiquetas:
   a_Array(1, 0) = "Código"
   a_Array(1, 1) = "Descripcion"
   a_Array(1, 2) = Empty
   a_Array(1, 3) = Empty
   'Tipo de Dato:
   a_Array(2, 0) = "C"
   a_Array(2, 1) = "C"
   a_Array(2, 2) = "C"
   a_Array(2, 3) = "D"
   'Ancho de campo:
   a_Array(3, 0) = 3
   a_Array(3, 1) = 40
   a_Array(3, 2) = 8
   a_Array(3, 3) = Empty
   'Campo Clave:
   a_Array(4, 0) = True
   a_Array(4, 1) = False
   a_Array(4, 2) = False
   a_Array(4, 3) = False
   'Valores Ingresados por el Sistema:
   a_Array(5, 0) = Empty
   a_Array(5, 1) = Empty
   a_Array(5, 2) = VGusuario
   a_Array(5, 3) = Date
   'a_Array(5, 4) = Format(Now, "aaaa-mm-dd hh:mm:ss.000")
   'Permite Nulos:
   a_Array(6, 0) = False
   a_Array(6, 1) = False
   a_Array(6, 2) = False
   a_Array(6, 3) = False
   
   mTablaBasica1.conexion = VGCNx
   mTablaBasica1.NombreTabla = "ct_tipoanalitico"
   mTablaBasica1.TituloForm = "Tipo de Analítico"
   mTablaBasica1.Filtro = "tipoanaliticocodigo<>'00'"
   mTablaBasica1.Arreglo = a_Array
   mTablaBasica1.Setear_Controles
   mTablaBasica1.Obtener_Campos
   mTablaBasica1.cargar_datos
   
   'oCrystalReport.ReportFileName = RutaRep & "MantMoneda.rpt"
   
End Sub

'FIXIT: Declare 'indice' con un tipo de datos de enlace en tiempo de compilación           FixIT90210ae-R1672-R1B8ZE
Private Sub mTablaBasica1_Click(indice As Variant)
' If indice = 3 Then
'   MDIPrincipal.cryRpt.Destination = crptToWindow
'   MDIPrincipal.cryRpt.WindowState = crptMaximized
'   MDIPrincipal.cryRpt.ReportFileName = App.Path & "\Reportes\rptTipoAnalitico.rpt"
'   MDIPrincipal.cryRpt.Connect = vgCADENAREPORT
'   MDIPrincipal.cryRpt.DiscardSavedData = True
'   MDIPrincipal.cryRpt.Action = 1
' End If
  If indice = 3 Then Call Impresion("rptTipoAnalitico.rpt")


End Sub
