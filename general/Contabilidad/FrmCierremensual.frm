VERSION 5.00
Object = "{272034D2-AC5F-11D6-810B-0050BAA91DB7}#18.0#0"; "mTablaBasica.ocx"
Begin VB.Form FrmCierremensual 
   Caption         =   "Cierre Mensual"
   ClientHeight    =   6975
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8475
   LinkTopic       =   "Form1"
   ScaleHeight     =   6975
   ScaleWidth      =   8475
   StartUpPosition =   2  'CenterScreen
   Begin MantTablaBasica.mTablaBasica mTablaBasica1 
      Height          =   6255
      Left            =   60
      TabIndex        =   0
      Top             =   90
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   11033
   End
End
Attribute VB_Name = "FrmCierremensual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim a_Array(0 To 12, 0 To 12)

Private Sub Form_Load()
      
   'Nombre Campos:
   a_Array(0, 0) = "empresacodigo"
   a_Array(0, 1) = "mes"
   a_Array(0, 2) = "anio"
   a_Array(0, 3) = "compras"
   a_Array(0, 4) = "inventarios"
   a_Array(0, 5) = "pagar"
   a_Array(0, 6) = "tesoreria"
   a_Array(0, 7) = "ventas"
   a_Array(0, 8) = "cobrar"
   a_Array(0, 9) = "contabilidad"

   'Etiquetas:
   a_Array(1, 0) = "Empresa" 'Empty
   a_Array(1, 1) = "Mes" 'Empty
   a_Array(1, 2) = "Año" 'Empty
   a_Array(1, 3) = "Compras"
   a_Array(1, 4) = "Inventarios"
   a_Array(1, 5) = "Pagar"
   a_Array(1, 6) = "Tesoreria"
   a_Array(1, 7) = "Ventas"
   a_Array(1, 8) = "Cobrar"
   a_Array(1, 9) = "Contabilidad"
   
   'Tipo de Dato:
   a_Array(2, 0) = "C"
   a_Array(2, 1) = "C"
   a_Array(2, 2) = "C"
   a_Array(2, 3) = "B"
   a_Array(2, 4) = "B"
   a_Array(2, 5) = "B"
   a_Array(2, 6) = "B"
   a_Array(2, 7) = "B"
   a_Array(2, 8) = "B"
   a_Array(2, 9) = "B"
   
   'Ancho de campo:
   a_Array(3, 0) = 2
   a_Array(3, 1) = 2
   a_Array(3, 2) = 4
   a_Array(3, 3) = 1
   a_Array(3, 4) = 1
   a_Array(3, 5) = 1
   a_Array(3, 6) = 1
   a_Array(3, 7) = 1
   a_Array(3, 8) = 1
   a_Array(3, 9) = 1
  
   'Campo Clave:
   a_Array(4, 0) = True
   a_Array(4, 1) = True
   a_Array(4, 2) = True
   a_Array(4, 3) = False
   a_Array(4, 4) = False
   a_Array(4, 5) = False
   a_Array(4, 6) = False
   a_Array(4, 7) = False
   a_Array(4, 8) = False
   a_Array(4, 9) = False
      
   'Valores Ingresados por el Sistema:
   a_Array(5, 0) = Empty 'VGParametros.empresacodigo
   a_Array(5, 1) = Empty 'VGParamSistem.Mesproceso
   a_Array(5, 2) = Empty 'VGParamSistem.Anoproceso
   a_Array(5, 3) = Empty
   a_Array(5, 4) = Empty
   a_Array(5, 5) = Empty
   a_Array(5, 6) = Empty
   a_Array(5, 7) = Empty
   a_Array(5, 8) = Empty
   a_Array(5, 9) = Empty
     
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
   a_Array(6, 9) = False
   
   
   mTablaBasica1.conexion = VGCNx
   mTablaBasica1.NombreTabla = "ct_cierremensual"
   mTablaBasica1.TituloForm = "Cierres de Sistema"
   mTablaBasica1.Arreglo = a_Array
'   mTablaBasica1.Filtro = "empresacodigo='" & VGParametros.empresacodigo & "' and anio='" & VGParamSistem.Anoproceso & "' "
   mTablaBasica1.Setear_Controles
   mTablaBasica1.Obtener_Campos
   mTablaBasica1.cargar_datos
   
   'oCrystalReport.ReportFileName = RutaRep & "MantMoneda.rpt"
   
End Sub

Private Sub mTablaBasica1_Click(indice As Variant)
' If indice = 3 Then
'   cryRpt.Destination = crptToWindow
'   cryRpt.WindowState = crptMaximized
'   cryRpt.ReportFileName = App.Path & "\Reportes\rptAplicacion.rpt"
'   cryRpt.Connect = vgCADENAREPORT
'   cryRpt.DiscardSavedData = True
'   cryRpt.Action = 1
' End If
  If indice = 3 Then Call Impresion("rptAplicacion.rpt")

End Sub

