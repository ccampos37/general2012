VERSION 5.00
Object = "{FED6C0D4-BBAF-48FE-B6CE-FFC87978CBAE}#215.0#0"; "Controles.ocx"
Begin VB.Form FrmCodigocajas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Codigo de Caja"
   ClientHeight    =   8910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9540
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8910
   ScaleWidth      =   9540
   StartUpPosition =   3  'Windows Default
   Begin UMantenimiento.TablasBasicas oTablasBasicas 
      Height          =   8835
      Left            =   210
      TabIndex        =   0
      Top             =   30
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   15584
   End
End
Attribute VB_Name = "FrmCodigocajas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim a_Array(0 To 12, 0 To 13)

Private Sub Form_Load()
   MostrarForm Me, "C2"
        
   'Nombre Campos:
   a_Array(0, 0) = "cajacodigo"
   a_Array(0, 1) = "cajadescripcion"
   a_Array(0, 2) = "cajadesccorta"
   a_Array(0, 3) = "cajacuentasoles"
   a_Array(0, 4) = "cajacuentadolares"
   a_Array(0, 5) = "cajaRendiciones"
   a_Array(0, 6) = "rendicionnumero01"
   a_Array(0, 7) = "rendicionnumero02"
   a_Array(0, 8) = "CajaCuentaxRendir"
   a_Array(0, 9) = "Cajasuspendida"
   a_Array(0, 10) = "CajaFondofijo"
   a_Array(0, 11) = "usuariocodigo"
   a_Array(0, 12) = "fechaact"
   'Etiquetas:
   a_Array(1, 0) = "Código"
   a_Array(1, 1) = "Descripcion"
   a_Array(1, 2) = "Descripción Corta"
   a_Array(1, 3) = "Cuenta Ctble Soles"
   a_Array(1, 4) = "Cuenta Ctble Dolares"
   a_Array(1, 5) = "Administra Rendiciones"
   a_Array(1, 6) = "Nro Rendicion Soles"
   a_Array(1, 7) = "Nro Rendicion Dolares"
   a_Array(1, 8) = "Admnistra Doc.x Rendir"
   a_Array(1, 9) = "Caja Suspendida"
   a_Array(1, 10) = "Caja Fondo Fijo"
   a_Array(1, 11) = ""
   a_Array(1, 12) = ""
   
   'Tipo de Dato:
   
   a_Array(2, 0) = "C"
   a_Array(2, 1) = "C"
   a_Array(2, 2) = "C"
   a_Array(2, 3) = "C"
   a_Array(2, 4) = "C"
   a_Array(2, 5) = "B"
   a_Array(2, 6) = "C"
   a_Array(2, 7) = "C"
   a_Array(2, 8) = "B"
   a_Array(2, 9) = "B"
   a_Array(2, 10) = "B"
   a_Array(2, 11) = "C"
   a_Array(2, 12) = "D"
   
   'Ancho de campo:
   
   a_Array(3, 0) = 2
   a_Array(3, 1) = 35
   a_Array(3, 2) = 20
   a_Array(3, 3) = 7
   a_Array(3, 4) = 7
   a_Array(3, 5) = 1
   a_Array(3, 6) = 6
   a_Array(3, 7) = 6
   a_Array(3, 8) = 1
   a_Array(3, 9) = 1
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
   a_Array(5, 11) = VGusuario
   a_Array(5, 12) = Date
   
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
   a_Array(6, 8) = True
   a_Array(6, 9) = True
   a_Array(6, 10) = True
   a_Array(6, 11) = True
   a_Array(6, 12) = True
   
   oTablasBasicas.conexion = VGCNx
   oTablasBasicas.NombreTabla = "te_codigocaja"
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
        Call Imprimir("Reptecodigocaja.rpt")
   End If
End Sub
Private Sub oTablasBasicas_txtCodigoLostFocus(indice2 As Variant)  'Formatea con ceros el campo codigo
    If indice2 = 0 Then
        Call oTablasBasicas.Formatear_Codigo(indice2)
    End If
End Sub

 

Private Sub TablasBasicas1_Click(indice As Variant)

End Sub

