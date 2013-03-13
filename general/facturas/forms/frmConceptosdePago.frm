VERSION 5.00
Object = "{FED6C0D4-BBAF-48FE-B6CE-FFC87978CBAE}#215.0#0"; "Controles.ocx"
Begin VB.Form frmConceptosdePago 
   Caption         =   "Conceptos de Pago"
   ClientHeight    =   8985
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9510
   LinkTopic       =   "Form1"
   ScaleHeight     =   8985
   ScaleWidth      =   9510
   StartUpPosition =   3  'Windows Default
   Begin UMantenimiento.TablasBasicas TablasBasicas1 
      Height          =   8775
      Left            =   240
      TabIndex        =   0
      Top             =   0
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   15478
   End
End
Attribute VB_Name = "frmConceptosdePago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim a_Array(0 To 12, 0 To 12)
Private Sub Form_Load()

   MostrarFormVentas Me, "C2"
   'Nombre Campos:
   a_Array(0, 0) = "pagocodigo"
   a_Array(0, 1) = "pagodescripcion"
   a_Array(0, 2) = "pagoefectivo"
   
   'Etiquetas:
   a_Array(1, 0) = "Código"
   a_Array(1, 1) = "Nombres"
   a_Array(1, 2) = "Pago efectivo(1) "
   'Tipo de Dato:
   a_Array(2, 0) = "C"
   a_Array(2, 1) = "C"
   a_Array(2, 2) = "N"
   'Ancho de campo:
   a_Array(3, 0) = 2
   a_Array(3, 1) = 30
   a_Array(3, 2) = 1
   'Campo Clave:
   a_Array(4, 0) = True
   a_Array(4, 1) = False
   a_Array(4, 2) = False
   'Valores Ingresados por el Sistema:
   a_Array(5, 0) = ""
   a_Array(5, 1) = ""
   a_Array(5, 2) = ""
   'a_Array(5, 4) = Format(Now, "aaaa-mm-dd hh:mm:ss.000")
   'Permite Nulos:
   a_Array(6, 0) = False
   a_Array(6, 1) = False
   a_Array(6, 2) = False
   
   TablasBasicas1.conexion = VGCNx
   TablasBasicas1.NombreTabla = "vt_conceptosdepago"
   TablasBasicas1.Arreglo = a_Array
   TablasBasicas1.Setear_Controles
   TablasBasicas1.Obtener_Campos
   TablasBasicas1.cargar_datos
   
End Sub

Private Sub oTablasBasicas_Click(indice As Variant)
   If indice = 3 Then
        Call imprimir("RepvtMantVend.rpt")
    ElseIf indice = 0 Then
      TablasBasicas1.Estado_Default (10)
   End If
End Sub
Private Sub oTablasBasicas_txtCodigoLostFocus(indice2 As Variant)  'Formatea con ceros el campo codigo
    If indice2 = 0 Then
        Call TablasBasicas1.Formatear_Codigo(indice2)
    End If
End Sub

