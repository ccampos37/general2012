VERSION 5.00
Object = "{FED6C0D4-BBAF-48FE-B6CE-FFC87978CBAE}#215.0#0"; "Controles.ocx"
Begin VB.Form FrmPuntoVenta 
   Caption         =   "Form1"
   ClientHeight    =   8790
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8940
   LinkTopic       =   "Form1"
   ScaleHeight     =   8790
   ScaleWidth      =   8940
   StartUpPosition =   3  'Windows Default
   Begin UMantenimiento.TablasBasicas oTablasBasicas 
      Height          =   8565
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   8625
      _ExtentX        =   15214
      _ExtentY        =   15108
   End
End
Attribute VB_Name = "FrmPuntoVenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim a_Array(0 To 12, 0 To 12)

Private Sub Form_Load()
   MostrarFormVentas Me, "C2"
        
   'Nombre Campos:
   a_Array(0, 0) = "puntovtacodigo"
   a_Array(0, 1) = "puntovtadescripcion"
   a_Array(0, 2) = "puntovtanropedido"
   a_Array(0, 3) = "puntovtanroguia"
   a_Array(0, 4) = "puntovtanrofact"
   a_Array(0, 5) = "puntovtanroguiarem"
   a_Array(0, 6) = "puntovtanotaabono"
   a_Array(0, 7) = "puntovtanotacargo"
   a_Array(0, 8) = "puntovtaautomat"
   a_Array(0, 9) = "puntovtaticket"
   a_Array(0, 10) = "codigocajavtas"
   a_Array(0, 11) = "usuariocodigo"
   
   'Etiquetas:
   a_Array(1, 0) = "Código"
   a_Array(1, 1) = "Descripcion"
   a_Array(1, 2) = "Nro.Pedido"
   a_Array(1, 3) = "Nro.Guia"
   a_Array(1, 4) = "Nro.Factura"
   a_Array(1, 5) = "Nro.Guia Remisión"
   a_Array(1, 6) = "Nota Abono"
   a_Array(1, 7) = "Nota Cargo"
   a_Array(1, 8) = "Automatizado"
   a_Array(1, 9) = "Ticket"
   a_Array(1, 10) = "Codigo Caja"
   a_Array(1, 11) = ""
   
   'Tipo de Dato:
   a_Array(2, 0) = "C"
   a_Array(2, 1) = "C"
   a_Array(2, 2) = "B"
   a_Array(2, 3) = "B"
   a_Array(2, 4) = "B"
   a_Array(2, 5) = "B"
   a_Array(2, 6) = "B"
   a_Array(2, 7) = "B"
   a_Array(2, 8) = "B"
   a_Array(2, 9) = "B"
   a_Array(2, 10) = "C"
   a_Array(2, 11) = "C"
   'Ancho de campo:
   a_Array(3, 0) = 2
   a_Array(3, 1) = 20
   a_Array(3, 2) = 1
   a_Array(3, 3) = 1
   a_Array(3, 4) = 1
   a_Array(3, 5) = 1
   a_Array(3, 6) = 1
   a_Array(3, 7) = 1
   a_Array(3, 8) = 1
   a_Array(3, 9) = 1
   a_Array(3, 10) = 2
   a_Array(3, 11) = 8
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
   'a_Array(5, 4) = Format(Now, "aaaa-mm-dd hh:mm:ss.000")
   'Permite Nulos:
   a_Array(6, 0) = False
   a_Array(6, 1) = False
   a_Array(6, 2) = True
   a_Array(6, 3) = False
   a_Array(6, 4) = True
   a_Array(6, 5) = True
   a_Array(6, 6) = False
   a_Array(6, 7) = True
   a_Array(6, 8) = True
   a_Array(6, 9) = True
   a_Array(6, 10) = True
   a_Array(6, 11) = False

   
   oTablasBasicas.conexion = VGCNx
   oTablasBasicas.NombreTabla = "vt_puntoventa"
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
        Call Imprimir("RepvtPuntoVta.rpt")
   End If
End Sub
Private Sub oTablasBasicas_txtCodigoLostFocus(indice2 As Variant)  'Formatea con ceros el campo codigo
    If indice2 = 0 Then
        Call oTablasBasicas.Formatear_Codigo(indice2)
    End If
End Sub
