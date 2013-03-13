VERSION 5.00
Object = "{FED6C0D4-BBAF-48FE-B6CE-FFC87978CBAE}#207.0#0"; "Controles.ocx"
Begin VB.Form FrmTransaccion 
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
Attribute VB_Name = "FrmTransaccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim a_Array(0 To 12, 0 To 12)

Private Sub Form_Load()
   MostrarForm Me, "C"
        
   'Nombre Campos:
   a_Array(0, 0) = "transaccioncodigo"
   a_Array(0, 1) = "transacciondescripcion"
   a_Array(0, 2) = "transaccionautomat"
   a_Array(0, 3) = "transaccioningsal"
   a_Array(0, 4) = "transaccionorigen"
   a_Array(0, 5) = "usuariocodigo"
   a_Array(0, 6) = "fechaact"
   'Etiquetas:
   a_Array(1, 0) = "C�digo"
   a_Array(1, 1) = "Descripcion"
   a_Array(1, 2) = "Automatizado"
   a_Array(1, 3) = "Ingreso Sal."
   a_Array(1, 4) = "Origen"
   a_Array(1, 5) = ""
   a_Array(1, 6) = ""
   'Tipo de Dato:
   a_Array(2, 0) = "C"
   a_Array(2, 1) = "C"
   a_Array(2, 2) = "B"
   a_Array(2, 3) = "B"
   a_Array(2, 4) = "C"
   a_Array(2, 5) = "C"
   a_Array(2, 6) = "D"
   'Ancho de campo:
   a_Array(3, 0) = 6
   a_Array(3, 1) = 20
   a_Array(3, 2) = 1
   a_Array(3, 3) = 1
   a_Array(3, 4) = 20
   a_Array(3, 5) = 8
   a_Array(3, 6) = ""
   'Campo Clave:
   a_Array(4, 0) = True
   a_Array(4, 1) = False
   a_Array(4, 2) = False
   a_Array(4, 3) = False
   a_Array(4, 4) = False
   a_Array(4, 5) = False
   a_Array(4, 6) = False
   'Valores Ingresados por el Sistema:
   a_Array(5, 0) = ""
   a_Array(5, 1) = ""
   a_Array(5, 2) = ""
   a_Array(5, 3) = ""
   a_Array(5, 4) = ""
   a_Array(5, 5) = g_usuario
   a_Array(5, 6) = Date
   'a_Array(5, 4) = Format(Now, "aaaa-mm-dd hh:mm:ss.000")
   'Permite Nulos:
   a_Array(6, 0) = False
   a_Array(6, 1) = False
   a_Array(6, 2) = True
   a_Array(6, 3) = True
   a_Array(6, 4) = True
   a_Array(6, 5) = False
   a_Array(6, 6) = False
   
   oTablasBasicas.conexion = cn
   oTablasBasicas.NombreTabla = "vt_transaccion"
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
        Call Imprimir("MantTransaccion.rpt")
   End If
End Sub
Private Sub oTablasBasicas_txtCodigoLostFocus(indice2 As Variant)  'Formatea con ceros el campo codigo
    If indice2 = 0 Then
        Call oTablasBasicas.Formatear_Codigo(indice2)
    End If
End Sub
