VERSION 5.00
Object = "{FED6C0D4-BBAF-48FE-B6CE-FFC87978CBAE}#215.0#0"; "Controles.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form FrmUnidad 
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
   Begin Crystal.CrystalReport oCrystalReport 
      Left            =   360
      Top             =   7920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "FrmUnidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim a_Array(0 To 12, 0 To 12)


Private Sub Form_Load()
   MostrarForm Me, "C2"
        
   'Nombre Campos:
   a_Array(0, 0) = "unidadcodigo"
   a_Array(0, 1) = "unidaddescripcion"
   a_Array(0, 2) = "unidaddescrcorta"
   a_Array(0, 3) = "unidadreferencial"
   a_Array(0, 4) = "unidadfactorconv"
   a_Array(0, 5) = "estadoreg"
   a_Array(0, 6) = "usuariocodigo"
   a_Array(0, 7) = "fechaact"
   'Etiquetas:
   a_Array(1, 0) = "Código"
   a_Array(1, 1) = "Descripcion"
   a_Array(1, 2) = "Descripción Corta"
   a_Array(1, 3) = "Unidad Referencial"
   a_Array(1, 4) = "Factor de Conversión"
   a_Array(1, 5) = "Activo"
   a_Array(1, 6) = ""
   a_Array(1, 7) = ""
   'Tipo de Dato:
   a_Array(2, 0) = "C"
   a_Array(2, 1) = "C"
   a_Array(2, 2) = "C"
   a_Array(2, 3) = "C"
   a_Array(2, 4) = "N"
   a_Array(2, 5) = "B"
   a_Array(2, 6) = "C"
   a_Array(2, 7) = "D"
   'Ancho de campo:
   a_Array(3, 0) = 3
   a_Array(3, 1) = 30
   a_Array(3, 2) = 15
   a_Array(3, 3) = 3
   a_Array(3, 4) = 8
   a_Array(3, 5) = 1
   a_Array(3, 6) = 8
   a_Array(3, 7) = ""
   'Campo Clave:
   a_Array(4, 0) = True
   a_Array(4, 1) = False
   a_Array(4, 2) = False
   a_Array(4, 3) = False
   a_Array(4, 4) = False
   a_Array(4, 5) = False
   a_Array(4, 6) = False
   a_Array(4, 7) = False
   'Valores Ingresados por el Sistema:
   a_Array(5, 0) = ""
   a_Array(5, 1) = ""
   a_Array(5, 2) = ""
   a_Array(5, 3) = ""
   a_Array(5, 4) = ""
   a_Array(5, 5) = ""
   a_Array(5, 6) = g_usuario
   a_Array(5, 7) = Date
   'a_Array(5, 4) = Format(Now, "aaaa-mm-dd hh:mm:ss.000")
   'a_Array(5, 4) = Format(Now, "ddddd hhhhh")
   'Permite Nulos:
   a_Array(6, 0) = False
   a_Array(6, 1) = False
   a_Array(6, 2) = False
   a_Array(6, 3) = True
   a_Array(6, 4) = True
   a_Array(6, 5) = True
   a_Array(6, 6) = False
   a_Array(6, 7) = False
   
   oTablasBasicas.Conexion = VGCNx
   oTablasBasicas.NombreTabla = "vt_unidad"
   oTablasBasicas.Arreglo = a_Array
   oTablasBasicas.Setear_Controles
   oTablasBasicas.Obtener_Campos
   oTablasBasicas.cargar_datos
   
   'oCrystalReport.ReportFileName = VGParamSistem.Rutareport & "MantUnidad.rpt"
   'oCrystalReport.Connect = conexion
   
End Sub

Private Sub oTablasBasicas_Click(indice As Variant)
   If indice = 3 Then
     'oCrystalReport.PrinterSelect
     oCrystalReport.DiscardSavedData = True
     oCrystalReport.Action = 1
   ElseIf indice = 0 Then
      oTablasBasicas.Estado_Default (5)
   End If
End Sub
Private Sub oTablasBasicas_txtCodigoLostFocus(indice2 As Variant)  'Formatea con ceros el campo codigo
    If indice2 = 0 Then
        Call oTablasBasicas.Formatear_Codigo(indice2)
    End If
End Sub
