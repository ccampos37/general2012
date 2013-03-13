VERSION 5.00
Object = "{FED6C0D4-BBAF-48FE-B6CE-FFC87978CBAE}#215.0#0"; "Controles.ocx"
Begin VB.Form FrmVendedor 
   Caption         =   "Form1"
   ClientHeight    =   8790
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8910
   LinkTopic       =   "Form1"
   ScaleHeight     =   8790
   ScaleWidth      =   8910
   StartUpPosition =   3  'Windows Default
   Begin UMantenimiento.TablasBasicas oTablasBasicas 
      Height          =   8895
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   15690
   End
End
Attribute VB_Name = "FrmVendedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim a_Array(0 To 12, 0 To 12)

Private Sub Form_Load()
   MostrarForm Me, "C2"
        
   'Nombre Campos:
   a_Array(0, 0) = "vendedorcodigo"
   a_Array(0, 1) = "vendedornombres"
   a_Array(0, 2) = "vendedordireccion"
   a_Array(0, 3) = "vendedorruc"
   a_Array(0, 4) = "vendedortelefono"
   a_Array(0, 5) = "vendedorreferencia"
   a_Array(0, 6) = "vendedorle"
   a_Array(0, 7) = "vendedorcomis1"
   a_Array(0, 8) = "vendedorcomis2"
   a_Array(0, 9) = "vendedorcomis3"
   a_Array(0, 10) = "estadoreg"
   a_Array(0, 11) = "usuariocodigo"
   a_Array(0, 12) = "fechaact"
   'Etiquetas:
   a_Array(1, 0) = "Código"
   a_Array(1, 1) = "Nombres"
   a_Array(1, 2) = "Dirección"
   a_Array(1, 3) = "RUC"
   a_Array(1, 4) = "Teléfono"
   a_Array(1, 5) = "Referencia"
   a_Array(1, 6) = "D.N.I."
   a_Array(1, 7) = "Comisión 1"
   a_Array(1, 8) = "Comisión 2"
   a_Array(1, 9) = "Comisión 3"
   a_Array(1, 10) = "Activo"
   a_Array(1, 11) = ""
   a_Array(1, 12) = ""
   'Tipo de Dato:
   a_Array(2, 0) = "C"
   a_Array(2, 1) = "C"
   a_Array(2, 2) = "C"
   a_Array(2, 3) = "C"
   a_Array(2, 4) = "C"
   a_Array(2, 5) = "C"
   a_Array(2, 6) = "C"
   a_Array(2, 7) = "N"
   a_Array(2, 8) = "N"
   a_Array(2, 9) = "N"
   a_Array(2, 10) = "B"
   a_Array(2, 11) = "C"
   a_Array(2, 12) = "D"
   'Ancho de campo:
   a_Array(3, 0) = 3
   a_Array(3, 1) = 50
   a_Array(3, 2) = 30
   a_Array(3, 3) = 15
   a_Array(3, 4) = 15
   a_Array(3, 5) = 30
   a_Array(3, 6) = 8
   a_Array(3, 7) = 8
   a_Array(3, 8) = 8
   a_Array(3, 9) = 8
   a_Array(3, 10) = 1
   a_Array(3, 11) = 8
   a_Array(3, 12) = ""
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
   a_Array(4, 12) = False
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
   a_Array(5, 9) = ""
   a_Array(5, 10) = ""
   a_Array(5, 11) = g_usuario
   a_Array(5, 12) = Date
   'a_Array(5, 4) = Format(Now, "aaaa-mm-dd hh:mm:ss.000")
   'Permite Nulos:
   a_Array(6, 0) = False
   a_Array(6, 1) = False
   a_Array(6, 2) = False
   a_Array(6, 3) = True
   a_Array(6, 4) = True
   a_Array(6, 5) = True
   a_Array(6, 6) = False
   a_Array(6, 7) = True
   a_Array(6, 8) = True
   a_Array(6, 9) = True
   a_Array(6, 10) = True
   a_Array(6, 11) = False
   a_Array(6, 12) = False
   
   oTablasBasicas.conexion = VGCNx
   oTablasBasicas.NombreTabla = "vt_vendedor"
   oTablasBasicas.Arreglo = a_Array
   oTablasBasicas.Setear_Controles
   oTablasBasicas.Obtener_Campos
   oTablasBasicas.cargar_datos
   
End Sub

Private Sub oTablasBasicas_Click(indice As Variant)
   If indice = 3 Then
        Call Imprimir("RepvtMantVend.rpt")
    ElseIf indice = 0 Then
      oTablasBasicas.Estado_Default (10)
   End If
End Sub
Private Sub oTablasBasicas_txtCodigoLostFocus(indice2 As Variant)  'Formatea con ceros el campo codigo
    If indice2 = 0 Then
        Call oTablasBasicas.Formatear_Codigo(indice2)
    End If
End Sub
