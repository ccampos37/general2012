VERSION 5.00
Object = "{FED6C0D4-BBAF-48FE-B6CE-FFC87978CBAE}#215.0#0"; "Controles.ocx"
Begin VB.Form FrmTipoConcepto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tipo de Conceptos"
   ClientHeight    =   8865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9600
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   9600
   StartUpPosition =   3  'Windows Default
   Begin UMantenimiento.TablasBasicas oTablasBasicas 
      Height          =   8805
      Left            =   270
      TabIndex        =   0
      Top             =   30
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   15531
   End
End
Attribute VB_Name = "FrmTipoConcepto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim a_Array(0 To 12, 0 To 12)

Private Sub Form_Load()
   MostrarForm Me, "C2"
        
   'Nombre Campos:
   a_Array(0, 0) = "Conceptocodigo"
   a_Array(0, 1) = "Conceptodescripcion"
   a_Array(0, 2) = "Conceptodesccorta"
   a_Array(0, 3) = "Conceptotipo"
   a_Array(0, 4) = "Conceptocuentasoles"
   a_Array(0, 5) = "Conceptocuentadolares"
   a_Array(0, 6) = "usuariocodigo"
   a_Array(0, 7) = "fechaact"
   
   'Etiquetas:
   a_Array(1, 0) = "Código"
   a_Array(1, 1) = "Descripcion"
   a_Array(1, 2) = "Desc. Corta"
   a_Array(1, 3) = "Cargo/Abono"
   a_Array(1, 4) = "Cta. Contable Soles"
   a_Array(1, 5) = "Cta. Contable Dolares"
   a_Array(1, 6) = ""
   a_Array(1, 7) = ""
   
   'Tipo de Dato:
   a_Array(2, 0) = "C"
   a_Array(2, 1) = "C"
   a_Array(2, 2) = "C"
   a_Array(2, 3) = "B"
   a_Array(2, 4) = "C"
   a_Array(2, 5) = "C"
   a_Array(2, 6) = "C"
   a_Array(2, 7) = "D"
   
   'Ancho de campo:
   a_Array(3, 0) = 2
   a_Array(3, 1) = 50
   a_Array(3, 2) = 30
   a_Array(3, 3) = 1
   a_Array(3, 4) = 20
   a_Array(3, 5) = 20
   a_Array(3, 6) = 8
   a_Array(3, 7) = ""
   
   'Campo Clave:
   a_Array(4, 0) = True
   a_Array(4, 1) = False
   a_Array(4, 2) = False
   a_Array(4, 3) = False
   a_Array(4, 4) = False
   a_Array(4, 5) = False
   
   'Valores Ingresados por el Sistema:
   a_Array(5, 0) = ""
   a_Array(5, 1) = ""
   a_Array(5, 2) = ""
   a_Array(5, 3) = ""
   a_Array(5, 4) = ""
   a_Array(5, 5) = ""
   a_Array(5, 6) = VGusuario
   a_Array(5, 7) = Date
   'a_Array(5, 4) = Format(Now, "aaaa-mm-dd hh:mm:ss.000")
   'Permite Nulos:
   a_Array(6, 0) = False
   a_Array(6, 1) = False
   a_Array(6, 2) = False
   a_Array(6, 3) = True
   a_Array(6, 4) = True
   a_Array(6, 5) = True
   a_Array(6, 6) = False
   a_Array(6, 7) = False
   
   oTablasBasicas.conexion = VGCNx
   oTablasBasicas.NombreTabla = "cp_conceptos"
   oTablasBasicas.Arreglo = a_Array
   oTablasBasicas.Setear_Controles
   oTablasBasicas.Obtener_Campos
   oTablasBasicas.cargar_datos
   
End Sub

Private Sub oTablasBasicas_Click(indice As Variant)
   If indice = 3 Then
        Call Imprimir("RepcpMantConcepto.rpt")
    ElseIf indice = 0 Then
      oTablasBasicas.Estado_Default (10)
   End If
End Sub
Private Sub oTablasBasicas_txtCodigoLostFocus(indice2 As Variant)  'Formatea con ceros el campo codigo
    If indice2 = 0 Then
        Call oTablasBasicas.Formatear_Codigo(indice2)
    End If
End Sub

