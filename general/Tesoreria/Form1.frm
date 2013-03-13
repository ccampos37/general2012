VERSION 5.00
Object = "{272034D2-AC5F-11D6-810B-0050BAA91DB7}#18.0#0"; "mtablabasica.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin MantTablaBasica.mTablaBasica mTablaBasica1 
      Height          =   6255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   11033
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim a_Array(0 To 12, 0 To 12)

Private Sub Form_Load()
   MostrarForm Me, "C2"
              
        
        
   'Nombre Campos:
   a_Array(0, 0) = "empresacodigo"
   a_Array(0, 1) = "empresadescripcion"
   a_Array(0, 2) = "empresadescrcorta"
   a_Array(0, 3) = "empresadireccion"
   a_Array(0, 4) = "empresaruc"
   a_Array(0, 5) = "empresatelefonos"
   a_Array(0, 6) = "operacionvalidacajabancos"
   a_Array(0, 7) = "operacioningresaobs"
   a_Array(0, 8) = "usuariocodigo"
   a_Array(0, 9) = "fechaact"
   'Etiquetas:
   a_Array(1, 0) = "Código"
   a_Array(1, 1) = "Descripcion"
   a_Array(1, 2) = "Descripción Corta"
   a_Array(1, 3) = "Direccion"
   a_Array(1, 4) = "Ruc"
   a_Array(1, 5) = "Telefonos"
   a_Array(1, 6) = "Valida Caja/Banco"
   a_Array(1, 7) = "Ingresa Observacion"
   a_Array(1, 8) = ""
   a_Array(1, 9) = ""
   
   'Tipo de Dato:
   
   a_Array(2, 0) = "C"
   a_Array(2, 1) = "C"
   a_Array(2, 2) = "C"
   a_Array(2, 3) = "C"
   a_Array(2, 4) = "C"
   a_Array(2, 5) = "C"
   a_Array(2, 6) = "B"
   a_Array(2, 7) = "B"
   a_Array(2, 8) = "C"
   a_Array(2, 9) = "D"
   
   'Ancho de campo:
   
   a_Array(3, 0) = 2
   a_Array(3, 1) = 35
   a_Array(3, 2) = 20
   a_Array(3, 3) = 20
   a_Array(3, 4) = 11
   a_Array(3, 5) = 20
   a_Array(3, 6) = 1
   a_Array(3, 7) = 1
   a_Array(3, 8) = 8
   a_Array(3, 9) = ""
   
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
   
   a_Array(5, 0) = ""
   a_Array(5, 1) = ""
   a_Array(5, 2) = ""
   a_Array(5, 3) = ""
   a_Array(5, 4) = ""
   a_Array(5, 5) = ""
   a_Array(5, 6) = ""
   a_Array(5, 7) = ""
   a_Array(5, 8) = g_usuario
   a_Array(5, 9) = Date
   
   'a_Array(5, 4) = Format(Now, "aaaa-mm-dd hh:mm:ss.000")
   'Permite Nulos:
   
   a_Array(6, 0) = False
   a_Array(6, 1) = False
   a_Array(6, 2) = False
   a_Array(6, 3) = True
   a_Array(6, 4) = False
   a_Array(6, 5) = True
   a_Array(6, 6) = False
   a_Array(6, 7) = False
   a_Array(6, 8) = False
   a_Array(6, 9) = False
   
   mTablasBasica1.Conexion = VGcnx
   mTablasBasica1.NombreTabla = "te_empresa"
   mTablasBasica1.Arreglo = a_Array
   mTablasBasica1.Setear_Controles
   mTablasBasica1.Obtener_Campos
   mTablasBasica1.cargar_datos
   
   ' Descripciones Duplicadas
   mTablasBasica1.DescripcionDuplicada = False
   mTablasBasica1.CampoDescripcion = 1

 End Sub

Private Sub mTablasBasica1_Click(indice As Variant)
   If indice = 3 Then
        Call Imprimir("Repteempresa.rpt")
   End If
End Sub
Private Sub mTablasBasica1_txtCodigoLostFocus(indice2 As Variant)  'Formatea con ceros el campo codigo
    If indice2 = 0 Then
        Call oTablasBasicas.Formatear_Codigo(indice2)
    End If
End Sub

 

