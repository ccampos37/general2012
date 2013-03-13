VERSION 5.00
Object = "{272034D2-AC5F-11D6-810B-0050BAA91DB7}#18.0#0"; "mTablaBasica.ocx"
Begin VB.Form FrmMntSolicitantes 
   Caption         =   "Form2"
   ClientHeight    =   6150
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7665
   LinkTopic       =   "Form2"
   ScaleHeight     =   6150
   ScaleWidth      =   7665
   StartUpPosition =   3  'Windows Default
   Begin MantTablaBasica.mTablaBasica mTablaBasica1 
      Height          =   6375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   11245
   End
End
Attribute VB_Name = "FrmMntSolicitantes"
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
   a_Array(0, 0) = "solicitantecodigo"
   a_Array(0, 1) = "solicitantenombre"
   'Etiquetas:
   a_Array(1, 0) = "Código"
   a_Array(1, 1) = "Nombre"
   'Tipo de Dato:
   a_Array(2, 0) = "C"
   a_Array(2, 1) = "C"
   'Ancho de campo:
   a_Array(3, 0) = 4
   a_Array(3, 1) = 20
   'Campo Clave:
   a_Array(4, 0) = True
   a_Array(4, 1) = False
   'Valores Ingresados por el Sistema:
   a_Array(5, 0) = Empty
   a_Array(5, 1) = Empty
   'a_Array(5, 4) = Format(Now, "aaaa-mm-dd hh:mm:ss.000")
   'Permite Nulos:
   a_Array(6, 0) = False
   a_Array(6, 1) = False
   
   mTablaBasica1.Conexion = VGCNx
   mTablaBasica1.NombreTabla = "co_solicitantes"
   mTablaBasica1.TituloForm = "Nombre de Solicitante"
   mTablaBasica1.Arreglo = a_Array
   mTablaBasica1.Setear_Controles
   mTablaBasica1.Obtener_Campos
   mTablaBasica1.cargar_datos
 
End Sub

Private Sub mTablaBasica1_Click(indice As Variant)
Dim arrform(2) As Variant
Dim arrparm(1) As Variant

If indice = 3 Then
    arrparm(0) = VGParamSistem.BDEmpresa
        
    arrform(0) = "@Empresa='" & VGparametros.NomEmpresa & "'"
    arrform(1) = "@ruc='" & VGparametros.RucEmpresa & "'"
    
    Call ImpresionRptProc("al_solicitantes.rpt", arrform, arrparm, , "Reporte de Solicitantes")
End If

End Sub



