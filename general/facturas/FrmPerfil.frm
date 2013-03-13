VERSION 5.00
Object = "{FED6C0D4-BBAF-48FE-B6CE-FFC87978CBAE}#215.0#0"; "Controles.ocx"
Begin VB.Form FrmPerfil 
   Caption         =   "Form1"
   ClientHeight    =   8820
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8925
   LinkTopic       =   "Form1"
   ScaleHeight     =   8820
   ScaleWidth      =   8925
   StartUpPosition =   3  'Windows Default
   Begin UMantenimiento.TablasBasicas oTablasBasicas 
      Height          =   8895
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   15690
   End
End
Attribute VB_Name = "FrmPerfil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim a_Array(0 To 12, 0 To 12)

Private Sub Form_Load()
   MostrarFormVentas Me, "C2"
        
   'Nombre Campos:
   a_Array(0, 0) = "perfilcodigo"
   a_Array(0, 1) = "perfildescripcion"
   a_Array(0, 2) = "usuariocodigo"
   a_Array(0, 3) = "fechaact"
   'Etiquetas:
   a_Array(1, 0) = "Código"
   a_Array(1, 1) = "Descripcion"
   a_Array(1, 2) = ""
   a_Array(1, 3) = ""
   'Tipo de Dato:
   a_Array(2, 0) = "C"
   a_Array(2, 1) = "C"
   a_Array(2, 2) = "C"
   a_Array(2, 3) = "D"
   'Ancho de campo:
   a_Array(3, 0) = 2
   a_Array(3, 1) = 20
   a_Array(3, 3) = 8
   a_Array(3, 4) = ""
   'Campo Clave:
   a_Array(4, 0) = True
   a_Array(4, 1) = False
   a_Array(4, 2) = False
   a_Array(4, 3) = False
   'Valores Ingresados por el Sistema:
   a_Array(5, 0) = ""
   a_Array(5, 1) = ""
   a_Array(5, 2) = g_usuario
   a_Array(5, 3) = Date
   'a_Array(5, 4) = Format(Now, "aaaa-mm-dd hh:mm:ss.000")
   'Permite Nulos:
   a_Array(6, 0) = False
   a_Array(6, 1) = True
   a_Array(6, 2) = False
   a_Array(6, 3) = False
   
   oTablasBasicas.conexion = VGCNx
   oTablasBasicas.NombreTabla = "gr_perfil"
   oTablasBasicas.Arreglo = a_Array
   oTablasBasicas.Setear_Controles
   oTablasBasicas.Obtener_Campos
   oTablasBasicas.cargar_datos
      ''''''''Descripciones Duplicadas
   oTablasBasicas.DescripcionDuplicada = False
   oTablasBasicas.CampoDescripcion = 1

End Sub

Private Sub oTablasBasicas_Click(indice As Variant)
   If indice = 3 Then
        Call imprimir("MantPerfil.rpt")
   End If
End Sub
Private Sub oTablasBasicas_txtCodigoLostFocus(indice2 As Variant)  'Formatea con ceros el campo codigo
    If indice2 = 0 Then
        Call oTablasBasicas.Formatear_Codigo(indice2)
    End If
End Sub
