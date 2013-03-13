VERSION 5.00
Object = "{FED6C0D4-BBAF-48FE-B6CE-FFC87978CBAE}#215.0#0"; "Controles.ocx"
Begin VB.Form FrmUsuario 
   Caption         =   "Form1"
   ClientHeight    =   8805
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8940
   LinkTopic       =   "Form1"
   ScaleHeight     =   8805
   ScaleWidth      =   8940
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
Attribute VB_Name = "FrmUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim a_Array(0 To 12, 0 To 12)

Private Sub Form_Load()
   MostrarForm Me, "C2"
        
   'Nombre Campos:
   a_Array(0, 0) = "usuariocodigo"
   a_Array(0, 1) = "usuarionombres"
   a_Array(0, 2) = "usuariopassword"
   a_Array(0, 3) = "usuariocias"
   a_Array(0, 4) = "usuarioopcion"
   a_Array(0, 5) = "usuarioopcform"
   a_Array(0, 6) = "estadoreg"
   a_Array(0, 7) = "usuarioingreso"
   a_Array(0, 8) = "fechaact"
   'Etiquetas:
   a_Array(1, 0) = "Código"
   a_Array(1, 1) = "Nombres"
   a_Array(1, 2) = "Password"
   a_Array(1, 3) = "Compañía"
   a_Array(1, 4) = "Opción"
   a_Array(1, 5) = "Opc.Form"
   a_Array(1, 6) = "Activo"
   a_Array(1, 7) = ""
   a_Array(1, 8) = ""
   'Tipo de Dato:
   a_Array(2, 0) = "C"
   a_Array(2, 1) = "C"
   a_Array(2, 2) = "C"
   a_Array(2, 3) = "C"
   a_Array(2, 4) = "C"
   a_Array(2, 5) = "C"
   a_Array(2, 6) = "B"
   a_Array(2, 7) = "C"
   a_Array(2, 8) = "D"
   'Ancho de campo:
   a_Array(3, 0) = 8
   a_Array(3, 1) = 30
   a_Array(3, 2) = 8
   a_Array(3, 3) = 1
   a_Array(3, 4) = 50
   a_Array(3, 5) = 50
   a_Array(3, 6) = 1
   a_Array(3, 7) = 8
   a_Array(3, 8) = ""
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
   a_Array(5, 0) = ""
   a_Array(5, 1) = ""
   a_Array(5, 2) = ""
   a_Array(5, 3) = ""
   a_Array(5, 4) = ""
   a_Array(5, 5) = ""
   a_Array(5, 6) = ""
   a_Array(5, 7) = g_usuario
   a_Array(5, 8) = Date
   'a_Array(5, 4) = Format(Now, "aaaa-mm-dd hh:mm:ss.000")
   'Permite Nulos:
   a_Array(6, 0) = False
   a_Array(6, 1) = False
   a_Array(6, 2) = False
   a_Array(6, 3) = True
   a_Array(6, 4) = True
   a_Array(6, 5) = True
   a_Array(6, 6) = True
   a_Array(6, 7) = False
   a_Array(6, 8) = False
   
   oTablasBasicas.conexion = VGcnx
   oTablasBasicas.NombreTabla = "gr_usuario"
   oTablasBasicas.Arreglo = a_Array
   oTablasBasicas.Setear_Controles
   oTablasBasicas.Obtener_Campos
   oTablasBasicas.cargar_datos
   
End Sub

Private Sub oTablasBasicas_Click(indice As Variant)
   If indice = 3 Then
        Call Imprimir("RepvtUsuario.rpt")
   ElseIf indice = 0 Then
      oTablasBasicas.Estado_Default (6)
   End If
End Sub
