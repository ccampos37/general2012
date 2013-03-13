VERSION 5.00
Object = "{272034D2-AC5F-11D6-810B-0050BAA91DB7}#18.0#0"; "mTablaBasica.ocx"
Begin VB.Form FrmMntEmpresas 
   Caption         =   "Form1"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7845
   LinkTopic       =   "Form1"
   ScaleHeight     =   6135
   ScaleWidth      =   7845
   StartUpPosition =   3  'Windows Default
   Begin MantTablaBasica.mTablaBasica mTablaBasica1 
      Height          =   6132
      Left            =   48
      TabIndex        =   0
      Top             =   48
      Width           =   7452
      _ExtentX        =   13150
      _ExtentY        =   10821
   End
End
Attribute VB_Name = "FrmMntEmpresas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim a_Array(0 To 12, 0 To 12)
Private Sub Form_Load()
   CentrarForm MDIPrincipal, Me
     
   'Nombre Campos:
   a_Array(0, 0) = "empresacodigo"
   a_Array(0, 1) = "empresadescripcion"
   a_Array(0, 2) = "agentederetencion"
   a_Array(0, 3) = "usuariocodigo"
   a_Array(0, 4) = "fechaact"
   'Etiquetas:
   a_Array(1, 0) = "Código"
   a_Array(1, 1) = "Descripcion"
   a_Array(1, 2) = "Agente de Retencion"
   a_Array(1, 3) = Empty
   a_Array(1, 4) = Empty
   'Tipo de Dato:
   a_Array(2, 0) = "C"
   a_Array(2, 1) = "C"
   a_Array(2, 2) = "C"
   a_Array(2, 3) = "C"
   a_Array(2, 4) = "D"
   'Ancho de campo:
   a_Array(3, 0) = 2
   a_Array(3, 1) = 20
   a_Array(3, 2) = 1
   a_Array(3, 3) = 8
   a_Array(3, 4) = Empty
   'Campo Clave:
   a_Array(4, 0) = True
   a_Array(4, 1) = False
   a_Array(4, 2) = False
   a_Array(4, 3) = False
   a_Array(4, 4) = False
   'Valores Ingresados por el Sistema:
   a_Array(5, 0) = Empty
   a_Array(5, 1) = Empty
   a_Array(5, 2) = Empty
   a_Array(5, 3) = VGusuario
   a_Array(5, 4) = Date
   'a_Array(5, 4) = Format(Now, "aaaa-mm-dd hh:mm:ss.000")
   'Permite Nulos:
   a_Array(6, 0) = False
   a_Array(6, 1) = False
   a_Array(6, 2) = False
   a_Array(6, 3) = False
   a_Array(6, 4) = False
   
   mTablaBasica1.Conexion = VGcnx
   mTablaBasica1.NombreTabla = "co_multiempresas"
   mTablaBasica1.TituloForm = "Tabla de Empresas"
    mTablaBasica1.Arreglo = a_Array
   mTablaBasica1.Setear_Controles
   mTablaBasica1.Obtener_Campos
   mTablaBasica1.cargar_datos
End Sub


