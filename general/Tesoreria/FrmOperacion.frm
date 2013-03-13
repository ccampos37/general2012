VERSION 5.00
Object = "{FED6C0D4-BBAF-48FE-B6CE-FFC87978CBAE}#215.0#0"; "Controles.ocx"
Begin VB.Form FrmOperacion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento Operacion General"
   ClientHeight    =   8700
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9285
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8700
   ScaleWidth      =   9285
   StartUpPosition =   2  'CenterScreen
   Begin UMantenimiento.TablasBasicas oTablasBasicas 
      Height          =   8775
      Left            =   150
      TabIndex        =   0
      Top             =   -60
      Width           =   9165
      _ExtentX        =   16166
      _ExtentY        =   15478
   End
End
Attribute VB_Name = "FrmOperacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim a_Array(0 To 12, 0 To 12)

Private Sub Form_Load()
   MostrarForm Me, "C2"
        
   'Nombre Campos:
   a_Array(0, 0) = "operacioncodigo"
   a_Array(0, 1) = "operaciondescripcion"
   a_Array(0, 2) = "operaciondesccorta"
   a_Array(0, 3) = "operacionmanejactas"
   a_Array(0, 4) = "operacioncontrolaclienteprov"
   a_Array(0, 5) = "operacioncontrolasaldos"
   a_Array(0, 6) = "operacionvalidacajabancos"
   a_Array(0, 7) = "operacioningresaobs"
   a_Array(0, 8) = "usuariocodigo"
   a_Array(0, 9) = "fechaact"
   'Etiquetas:
   a_Array(1, 0) = "Código"
   a_Array(1, 1) = "Descripcion"
   a_Array(1, 2) = "Descripción Corta"
   a_Array(1, 3) = "Maneja Cuenta"
   a_Array(1, 4) = "Operac.(C)liente/(P)rov/(X)"
   a_Array(1, 5) = "Control Saldos"
   a_Array(1, 6) = "Valida Caja/Banco"
   a_Array(1, 7) = "Observacion"
   a_Array(1, 8) = ""
   a_Array(1, 9) = ""
   'Tipo de Dato:
   a_Array(2, 0) = "C"
   a_Array(2, 1) = "C"
   a_Array(2, 2) = "C"
   a_Array(2, 3) = "B"
   a_Array(2, 4) = "C"
   a_Array(2, 5) = "B"
   a_Array(2, 6) = "C"
   a_Array(2, 7) = "B"
   
   a_Array(2, 8) = "C"
   a_Array(2, 9) = "D"
   'Ancho de campo:
   a_Array(3, 0) = 2
   a_Array(3, 1) = 35
   a_Array(3, 2) = 20
   a_Array(3, 3) = 1
   a_Array(3, 4) = 1
   a_Array(3, 5) = 1
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
   
   a_Array(5, 8) = VGusuario
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
   
   oTablasBasicas.Conexion = VGcnx
   oTablasBasicas.NombreTabla = "te_operaciongeneral"
   oTablasBasicas.Arreglo = a_Array
   oTablasBasicas.Setear_Controles
   oTablasBasicas.Obtener_Campos
   oTablasBasicas.cargar_datos
   
   ' Descripciones Duplicadas
   oTablasBasicas.DescripcionDuplicada = False
   oTablasBasicas.CampoDescripcion = 1

 End Sub

Private Sub oTablasBasicas_Click(indice As Variant)
   If indice = 3 Then
        Call Imprimir("Repteoperacion.rpt")
   End If
End Sub
Private Sub oTablasBasicas_txtCodigoLostFocus(indice2 As Variant)  'Formatea con ceros el campo codigo
    If indice2 = 0 Then
        Call oTablasBasicas.Formatear_Codigo(indice2)
    End If
End Sub

 
