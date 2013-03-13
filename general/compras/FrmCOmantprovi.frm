VERSION 5.00
Object = "{272034D2-AC5F-11D6-810B-0050BAA91DB7}#18.0#0"; "mTablaBasica.ocx"
Begin VB.Form FrmMntprovi 
   Caption         =   "Modo Provisiones"
   ClientHeight    =   6390
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7740
   LinkTopic       =   "Form1"
   ScaleHeight     =   6390
   ScaleWidth      =   7740
   StartUpPosition =   3  'Windows Default
   Begin MantTablaBasica.mTablaBasica mTablaBasica1 
      Height          =   6090
      Left            =   60
      TabIndex        =   0
      Top             =   105
      Width           =   7800
      _ExtentX        =   13758
      _ExtentY        =   10742
   End
End
Attribute VB_Name = "FrmMntprovi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim a_Array(0 To 12, 0 To 12)
Private Sub Form_Load()
   Me.Width = 7860: Me.Height = 6795
   mTablaBasica1.Width = 7590
   CentrarForm MDIPrincipal, Me
     
   'Nombre Campos:
   a_Array(0, 0) = "modoprovicod"
   a_Array(0, 1) = "modoprovidesc"
   a_Array(0, 2) = "modoprovictacte"
   a_Array(0, 3) = "modoproviregcom"
   a_Array(0, 4) = "modoprovitesor"
   a_Array(0, 5) = "modoprovireghon"
   a_Array(0, 6) = "modoproviobserv"
   a_Array(0, 7) = "modoproviflgcon"
   a_Array(0, 8) = "modoprovicontes"
   a_Array(0, 9) = "modoprovianalitico"
   a_Array(0, 10) = "usuariocodigo"
   a_Array(0, 11) = "fechaact"
   
   
   'Etiquetas:
   a_Array(1, 0) = "Código"
   a_Array(1, 1) = "Descripcion"
   a_Array(1, 2) = "Actualiza CtaCte"
   a_Array(1, 3) = "Emite RegCompras"
   a_Array(1, 4) = "Actualiza Tesoreria"
   a_Array(1, 5) = "Emite Reg.Honorarios"
   a_Array(1, 6) = "Ingresa Observacion"
   a_Array(1, 7) = "Contabiliza Provision"
   a_Array(1, 8) = "Contabiliza Tesoreria"
   a_Array(1, 9) = "Contabiliza Cuenta de Terceros"
   a_Array(1, 10) = Empty
   a_Array(1, 11) = Empty
   
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
   a_Array(2, 11) = "D"
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
   a_Array(3, 9) = 8
   a_Array(3, 10) = 8
   a_Array(3, 11) = Empty
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
   a_Array(5, 0) = Empty
   a_Array(5, 1) = Empty
   a_Array(5, 2) = Empty
   a_Array(5, 3) = Empty
   a_Array(5, 4) = Empty
   a_Array(5, 5) = Empty
   a_Array(5, 6) = Empty
   a_Array(5, 7) = Empty
   a_Array(5, 8) = Empty
   a_Array(5, 9) = Empty
   a_Array(5, 10) = VGUsuario
   a_Array(5, 11) = Date
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
   a_Array(6, 10) = False
   a_Array(6, 11) = False
   
   mTablaBasica1.Conexion = VGCNx
   mTablaBasica1.NombreTabla = "co_modoprovi"
   mTablaBasica1.TituloForm = "Modo de Provisión"
   'mTablaBasica1.Filtro = "tipoanaliticocodigo<>'00'"
   mTablaBasica1.Arreglo = a_Array
   mTablaBasica1.Setear_Controles
   mTablaBasica1.Obtener_Campos
   mTablaBasica1.cargar_datos
End Sub

