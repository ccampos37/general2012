VERSION 5.00
Object = "{FED6C0D4-BBAF-48FE-B6CE-FFC87978CBAE}#215.0#0"; "Controles.ocx"
Begin VB.Form FrmEmpresa 
   Caption         =   "Form1"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8925
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   8925
   StartUpPosition =   3  'Windows Default
   Begin UMantenimiento.TablasBasicas TablasBasicas1 
      Height          =   8655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   15266
   End
End
Attribute VB_Name = "FrmEmpresa"
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
   a_Array(0, 6) = "tdocumentocanje"
   a_Array(0, 7) = "tdocumentorenova"
   a_Array(0, 8) = "contabilizaenlinea"
   a_Array(0, 9) = "usuariocodigo"
   a_Array(0, 10) = "fechaact"
   'Etiquetas:
   a_Array(1, 0) = "Código"
   a_Array(1, 1) = "Descripcion"
   a_Array(1, 2) = "Descrip.Corta"
   a_Array(1, 3) = "Direccion"
   a_Array(1, 4) = "RUC"
   a_Array(1, 5) = "Telefonos"
   a_Array(1, 6) = "Doc.Canjes"
   a_Array(1, 7) = "Doc.Renovación"
   a_Array(1, 8) = "Contab.en Linea"
   
   a_Array(1, 9) = ""
   a_Array(1, 10) = ""
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
   
   a_Array(2, 9) = "C"
   a_Array(2, 10) = "D"
   'Ancho de campo:
   a_Array(3, 0) = 2
   a_Array(3, 1) = 30
   a_Array(3, 2) = 15
   a_Array(3, 3) = 30
   a_Array(3, 4) = 11
   a_Array(3, 5) = 20
   a_Array(3, 6) = 2
   a_Array(3, 7) = 2
   a_Array(3, 8) = 1
   
   a_Array(3, 9) = 8
   a_Array(3, 10) = ""
   'Campo Clave:
   a_Array(4, 0) = True
   a_Array(4, 1) = False
   a_Array(4, 2) = False
   a_Array(4, 3) = False
   a_Array(4, 4) = False
   a_Array(4, 5) = False
   a_Array(4, 6) = False
   a_Array(4, 7) = False
   
   a_Array(4, 9) = False
   a_Array(4, 10) = False
   'Valores Ingresados por el Sistema:
   a_Array(5, 0) = ""
   a_Array(5, 1) = ""
   a_Array(5, 2) = ""
   a_Array(5, 3) = ""
   a_Array(5, 4) = ""
   a_Array(5, 5) = ""
   a_Array(5, 6) = ""
   a_Array(5, 7) = ""
   a_Array(5, 8) = ""
  
   a_Array(5, 9) = VGusuario
   a_Array(5, 10) = Date
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
   
   a_Array(6, 9) = False
   a_Array(6, 10) = False
   
   TablasBasicas1.conexion = VGCNx
   TablasBasicas1.NombreTabla = "cp_parametros"
   TablasBasicas1.Arreglo = a_Array
   TablasBasicas1.Setear_Controles
   TablasBasicas1.Obtener_Campos
   TablasBasicas1.cargar_datos
      ''''''''Descripciones Duplicadas
   TablasBasicas1.DescripcionDuplicada = False
   TablasBasicas1.CampoDescripcion = 1

 End Sub

Private Sub oTablasBasicas_Click(indice As Variant)
   If indice = 3 Then
        Call Imprimir("RepcpMantEmp.rpt")
   End If
End Sub
Private Sub oTablasBasicas_txtCodigoLostFocus(indice2 As Variant)  'Formatea con ceros el campo codigo
    If indice2 = 0 Then
        Call TablasBasicas1.Formatear_Codigo(indice2)
    End If
End Sub
