VERSION 5.00
Object = "{FED6C0D4-BBAF-48FE-B6CE-FFC87978CBAE}#215.0#0"; "Controles.ocx"
Begin VB.Form FrmXXXCuentaBancaria 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9420
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8880
   ScaleWidth      =   9420
   StartUpPosition =   3  'Windows Default
   Begin UMantenimiento.TablasBasicas oTablasBasicas 
      Height          =   8775
      Left            =   210
      TabIndex        =   0
      Top             =   0
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   15478
   End
End
Attribute VB_Name = "FrmXXXCuentaBancaria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim a_Array(0 To 12, 0 To 12)

Private Sub Form_Load()
   MostrarForm Me, "C2"
        
   'Nombre Campos:
   a_Array(0, 0) = "cbanco_codigo"
   a_Array(0, 1) = "monedacodigo"
   a_Array(0, 2) = "cbanco_numero"
   a_Array(0, 3) = "cbanco_referenciacta"
   a_Array(0, 4) = "cbanco_nrocheque"
   a_Array(0, 5) = "cbanco_cuenta"
   a_Array(0, 6) = "cbanco_analitico"
   a_Array(0, 7) = "usuariocodigo"
   a_Array(0, 8) = "fechaact"
   
   'Etiquetas:
   a_Array(1, 0) = "Banco"
   a_Array(1, 1) = "Moneda"
   a_Array(1, 2) = "Cuenta Bancaria"
   a_Array(1, 3) = "Descripcion"
   a_Array(1, 4) = "No Inicio Cheque"
   a_Array(1, 5) = "Cuenta Contable"
   a_Array(1, 6) = "Cuenta Analitico"
   a_Array(1, 7) = ""
   a_Array(1, 8) = ""
   
   'Tipo de Dato:
   a_Array(2, 0) = "C"
   a_Array(2, 1) = "C"
   a_Array(2, 2) = "C"
   a_Array(2, 3) = "C"
   a_Array(2, 4) = "C"
   a_Array(2, 5) = "C"
   a_Array(2, 6) = "C"
   a_Array(2, 7) = "C"
   a_Array(2, 8) = "D"
   
   'Ancho de campo:
   a_Array(3, 0) = 2
   a_Array(3, 1) = 2
   a_Array(3, 2) = 20
   a_Array(3, 3) = 30
   a_Array(3, 4) = 15
   a_Array(3, 5) = 6
   a_Array(3, 6) = 11
   a_Array(3, 7) = 8
   a_Array(3, 8) = ""
   
   'Campo Clave:
   a_Array(4, 0) = True
   a_Array(4, 1) = False
   a_Array(4, 2) = True
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
   a_Array(5, 7) = VGusuario
   a_Array(5, 8) = Date
   
   'a_Array(5, 4) = Format(Now, "aaaa-mm-dd hh:mm:ss.000")
   'Permite Nulos:
   a_Array(6, 0) = False
   a_Array(6, 1) = False
   a_Array(6, 2) = False
   a_Array(6, 3) = False
   a_Array(6, 4) = False
   a_Array(6, 5) = False
   a_Array(6, 6) = False
   a_Array(6, 7) = False
   a_Array(6, 8) = False
   
   oTablasBasicas.Conexion = VGcnx
   oTablasBasicas.NombreTabla = "te_cuentabancos"
   oTablasBasicas.Arreglo = a_Array
   oTablasBasicas.Setear_Controles
   oTablasBasicas.Obtener_Campos
   oTablasBasicas.cargar_datos
      ''''''''Descripciones Duplicadas
   oTablasBasicas.DescripcionDuplicada = False
   'oTablasBasicas.CampoDescripcion = 1
   
End Sub

Private Sub oTablasBasicas_Click(indice As Variant)
   If indice = 3 Then
     Call Imprimir("RepvtMantBanc.rpt")
   End If
End Sub
Private Sub oTablasBasicas_txtCodigoLostFocus(indice2 As Variant)  'Formatea con ceros el campo codigo
    If indice2 = 0 Then
        Call oTablasBasicas.Formatear_Codigo(indice2)
    End If
End Sub

