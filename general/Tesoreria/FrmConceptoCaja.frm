VERSION 5.00
Object = "{FED6C0D4-BBAF-48FE-B6CE-FFC87978CBAE}#215.0#0"; "Controles.ocx"
Begin VB.Form FrmConceptocaja 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Bancos"
   ClientHeight    =   8865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9570
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   9570
   StartUpPosition =   3  'Windows Default
   Begin UMantenimiento.TablasBasicas oTablasBasicas 
      Height          =   8805
      Left            =   165
      TabIndex        =   0
      Top             =   60
      Width           =   9105
      _ExtentX        =   16060
      _ExtentY        =   15531
   End
End
Attribute VB_Name = "FrmConceptocaja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim a_Array(0 To 12, 0 To 12)

Private Sub Form_Load()
   MostrarForm Me, "C2"
   'Nombre Campos:
   a_Array(0, 0) = "conceptocodigo"
   a_Array(0, 1) = "conceptodescripcion"
   a_Array(0, 2) = "conceptodesccorta"
   a_Array(0, 3) = "conceptotipooperacion"
   a_Array(0, 4) = "conceptoingresaobs"
   a_Array(0, 5) = "conceptoingresaref"
   a_Array(0, 6) = "conceptoingresaref2"
   a_Array(0, 7) = "conceptocuentasoles"
   a_Array(0, 8) = "conceptocuentadolar"
   a_Array(0, 9) = "conceptosiccosto"
   a_Array(0, 10) = "conceptotextccosto"
   a_Array(0, 11) = "usuariocodigo"
   a_Array(0, 12) = "fechaact"
   
   
   'Etiquetas:
   a_Array(1, 0) = "Código"
   a_Array(1, 1) = "Descripcion"
   a_Array(1, 2) = "Descripción Corta"
   a_Array(1, 3) = "Tipo Operacion"
   a_Array(1, 4) = "Ingresa Observacion"
   a_Array(1, 5) = "Ingresa Entidad"
   a_Array(1, 6) = "Ingresa Ref.2"
   a_Array(1, 7) = "Cuenta Contable Soles"
   a_Array(1, 8) = "Cuenta Contable Dolar"
   a_Array(1, 9) = "Lleva Centro de Costos"
   a_Array(1, 10) = "Codigos de Centros de Costos"
   
   a_Array(1, 11) = Empty
   a_Array(1, 12) = Empty
   
   
   'Tipo de Dato:
   a_Array(2, 0) = "C"
   a_Array(2, 1) = "C"
   a_Array(2, 2) = "C"
   a_Array(2, 3) = "C"
   a_Array(2, 4) = "B"
   a_Array(2, 5) = "B"
   a_Array(2, 6) = "B"
   a_Array(2, 7) = "C"
   a_Array(2, 8) = "C"
   a_Array(2, 9) = "B"
   a_Array(2, 10) = "C"
   
   a_Array(2, 11) = "C"
   a_Array(2, 12) = "D"
   
   
   
   'Ancho de campo:
   a_Array(3, 0) = 2
   a_Array(3, 1) = 35
   a_Array(3, 2) = 20
   a_Array(3, 3) = 1
   a_Array(3, 4) = 1
   a_Array(3, 5) = 1
   a_Array(3, 6) = 1
   a_Array(3, 7) = 6
   a_Array(3, 8) = 6
   a_Array(3, 9) = 1
   a_Array(3, 10) = 100
   
   a_Array(3, 11) = 8
   a_Array(3, 12) = Empty
   
   
   
   
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
   a_Array(6, 3) = True
   a_Array(6, 4) = True
   a_Array(6, 5) = True
   a_Array(6, 6) = True
   a_Array(6, 7) = True
   a_Array(6, 8) = True
   a_Array(6, 9) = True
   a_Array(6, 10) = True
   
   a_Array(6, 11) = False
   a_Array(6, 12) = False
      
   
   oTablasBasicas.Conexion = VGcnx
   oTablasBasicas.NombreTabla = "te_conceptocaja"
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
        Call Imprimir("Repteconceptocaja.rpt")
   End If
End Sub
Private Sub oTablasBasicas_txtCodigoLostFocus(indice2 As Variant)  'Formatea con ceros el campo codigo
    If indice2 = 0 Then
        Call oTablasBasicas.Formatear_Codigo(indice2)
    End If
End Sub

 

Private Sub TablasBasicas1_Click(indice As Variant)

End Sub
