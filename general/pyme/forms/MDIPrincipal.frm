VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form MDIPrincipal 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Sistema de PyMe Integrado"
   ClientHeight    =   8655
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   14745
   FillColor       =   &H00E0E0E0&
   ForeColor       =   &H8000000A&
   Icon            =   "MDIPrincipal.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "MDIPrincipal.frx":15DBA
   ScaleHeight     =   8655
   ScaleWidth      =   14745
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6930
      Top             =   6525
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":160FC
            Key             =   "Entrar"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":28E05
            Key             =   "Retornar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":3B55A
            Key             =   "Camara"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":3BE87
            Key             =   "Tabla"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":3C6AF
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":3C7C1
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":3C8D3
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":3C9E5
            Key             =   "New"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   8280
      Width           =   14745
      _ExtentX        =   26009
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6174
            MinWidth        =   6174
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Crystal.CrystalReport CryRptProc 
      Left            =   5670
      Top             =   6660
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   7590
      Top             =   6480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":3CAF7
            Key             =   "Facturar"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":4608F
            Key             =   "Modificar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":554E8
            Key             =   "Nuevo"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":60149
            Key             =   "Eliminar"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":7345D
            Key             =   "Facturado"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":7DD9A
            Key             =   "Retornar"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":904EF
            Key             =   "Insertar"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":9A30C
            Key             =   "Sacar"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":A4186
            Key             =   "Grabar"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":B3746
            Key             =   "Adicionar"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":BE0DE
            Key             =   "Imprimir"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":D3CF2
            Key             =   "Consultar"
         EndProperty
      EndProperty
   End
   Begin Crystal.CrystalReport oCrystalReport 
      Left            =   6255
      Top             =   6705
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComctlLib.ImageList ImageList3 
      Left            =   0
      Top             =   1560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":E48F8
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":E4D4A
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":E4EA6
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":E500E
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":E516A
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":E52C6
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":E5422
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":E557E
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":E56DE
            Key             =   "IMG9"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar ToolComprob 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   14745
      _ExtentX        =   26009
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList3"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "k1"
            Description     =   "Nuevo"
            Object.ToolTipText     =   "Nuevo"
            ImageKey        =   "IMG4"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "k2"
            Description     =   "Grabar Salir"
            Object.ToolTipText     =   "Grabar y Salir"
            ImageKey        =   "IMG9"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "k3"
            Description     =   "Eliminar"
            Object.ToolTipText     =   "Eliminar"
            ImageKey        =   "IMG8"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "k4"
            Description     =   "Modificar"
            Object.ToolTipText     =   "Modificar"
            ImageKey        =   "IMG5"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "k5"
            Description     =   "Cancelar"
            Object.ToolTipText     =   "Cancelar Operacion"
            ImageKey        =   "IMG1"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "k6"
            Description     =   "Añadir Detalle"
            Object.ToolTipText     =   "Añadir Detalle"
            ImageKey        =   "IMG6"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "k7"
            Description     =   "Eliminar Detalle"
            Object.ToolTipText     =   "Eliminar Detalle"
            ImageKey        =   "IMG7"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "k8"
            Description     =   "Imprimir"
            Object.ToolTipText     =   "Imprimir"
            ImageKey        =   "IMG2"
         EndProperty
      EndProperty
   End
   Begin VB.Menu menu01 
      Caption         =   "Inventarios"
      Begin VB.Menu menu01_01 
         Caption         =   "Movimientos"
         Begin VB.Menu menu01_01_01 
            Caption         =   "Nota de Ingreso"
            Shortcut        =   ^I
         End
         Begin VB.Menu menu01_01_02 
            Caption         =   "Nota de Salidas "
            Shortcut        =   ^S
         End
         Begin VB.Menu menu01_01_03 
            Caption         =   "Traslados"
         End
         Begin VB.Menu menu01_01_04 
            Caption         =   "Ing. Requerimientos PEDIDOS"
         End
      End
      Begin VB.Menu menu01_02 
         Caption         =   "Tablas de Ayudas"
         Begin VB.Menu menu01_02_01 
            Caption         =   "&Artículos"
         End
         Begin VB.Menu menu01_02_02 
            Caption         =   "Al&macenes"
         End
         Begin VB.Menu menu01_02_03 
            Caption         =   "Tra&nsacciones"
         End
         Begin VB.Menu menu01_02_04 
            Caption         =   "Documentos"
         End
         Begin VB.Menu menu01_02_05 
            Caption         =   "Unidades de Medida"
         End
         Begin VB.Menu menu01_02_06 
            Caption         =   "Familia de Artículos"
         End
      End
      Begin VB.Menu menu01_03 
         Caption         =   "&Reportes"
         Begin VB.Menu menu01_03_01 
            Caption         =   "Católogo de Artículo"
         End
         Begin VB.Menu menu01_03_02 
            Caption         =   "Seguimiento de Requerimientos"
            Visible         =   0   'False
         End
         Begin VB.Menu menu01_03_03 
            Caption         =   "Stock de Artículos"
         End
         Begin VB.Menu menu01_03_04 
            Caption         =   "Kardex de Artículos"
         End
         Begin VB.Menu menu01_03_05 
            Caption         =   "Documentos"
            Begin VB.Menu menu01_03_05_01 
               Caption         =   "Detallado"
            End
            Begin VB.Menu menu01_03_05_02 
               Caption         =   "Resumido"
            End
            Begin VB.Menu menu01_03_05_03 
               Caption         =   "Informe de Traslados"
            End
         End
      End
      Begin VB.Menu menu01_04 
         Caption         =   "&Consultas"
         Begin VB.Menu menu01_04_01 
            Caption         =   "Saldos Consolidados x articulo"
         End
         Begin VB.Menu menu01_04_02 
            Caption         =   "Saldos Consolidados x Familia"
            Visible         =   0   'False
         End
         Begin VB.Menu menu01_04_03 
            Caption         =   "Saldos Consolidados"
         End
         Begin VB.Menu menu01_04_04 
            Caption         =   "Documentos"
         End
      End
      Begin VB.Menu menu01_05 
         Caption         =   "&Procesos"
         Begin VB.Menu menu01_05_01 
            Caption         =   "Recalculo Saldo Fisico"
         End
         Begin VB.Menu menu01_05_02 
            Caption         =   "Anulacion de Documentos"
            Begin VB.Menu menu01_05_02_01 
               Caption         =   "Documentos"
            End
            Begin VB.Menu menu01_05_02_02 
               Caption         =   "Transferencias"
            End
         End
      End
   End
   Begin VB.Menu menu02 
      Caption         =   "Ventas"
      Begin VB.Menu menu02_01 
         Caption         =   "Movimientos"
         Begin VB.Menu menu02_01_01 
            Caption         =   "Facturacion"
         End
         Begin VB.Menu menu02_01_02 
            Caption         =   "Correccion Documentos"
         End
         Begin VB.Menu menu02_01_03 
            Caption         =   "Correccion Documentos Forma de Pago"
         End
      End
      Begin VB.Menu menu02_02 
         Caption         =   "Consultas"
         Begin VB.Menu menu02_02_01 
            Caption         =   "Documentos"
         End
      End
      Begin VB.Menu menu02_03 
         Caption         =   "Reportes"
         Begin VB.Menu menu02_03_01 
            Caption         =   "Ventas detallado Punto Venta"
         End
         Begin VB.Menu menu02_03_02 
            Caption         =   "Ventas resumidasr Punto de venta"
         End
      End
      Begin VB.Menu menu02_04 
         Caption         =   "Tablas"
         Begin VB.Menu menu02_04_01 
            Caption         =   "Punto de ventas"
         End
         Begin VB.Menu menu02_04_02 
            Caption         =   "Punto Vta - Documento"
         End
         Begin VB.Menu menu02_04_03 
            Caption         =   "Modo de ventas"
         End
         Begin VB.Menu menu02_04_04 
            Caption         =   "Lista de Precios"
         End
         Begin VB.Menu menu02_04_05 
            Caption         =   "Clientes"
         End
      End
   End
   Begin VB.Menu menu03 
      Caption         =   "Cobranzas"
      Begin VB.Menu menu03_01 
         Caption         =   "Movimientos"
         Begin VB.Menu menu03_01_01 
            Caption         =   "Cargo de Interes"
         End
         Begin VB.Menu menu03_01_02 
            Caption         =   "Cobranza Clientes"
         End
         Begin VB.Menu menu03_01_03 
            Caption         =   "Gastos Varios"
         End
         Begin VB.Menu menu03_01_05 
            Caption         =   "Modifica Cargos Intereses"
            Visible         =   0   'False
         End
      End
      Begin VB.Menu menu03_03 
         Caption         =   "Consultas"
         Begin VB.Menu menu03_03_02 
            Caption         =   "Busqueda de Ckiente"
         End
      End
      Begin VB.Menu menu03_02 
         Caption         =   "Reportes"
         Begin VB.Menu menu03_02_01 
            Caption         =   "Saldos de Clientes"
         End
         Begin VB.Menu menu03_02_02 
            Caption         =   "Liquidacion Detallada"
         End
         Begin VB.Menu menu03_02_03 
            Caption         =   "Liquidacion Resumen"
         End
         Begin VB.Menu menu03_02_04 
            Caption         =   "Numero recibo"
         End
      End
      Begin VB.Menu menu03_04 
         Caption         =   "Tablas"
         Begin VB.Menu menu03_04_01 
            Caption         =   "Parametros"
         End
         Begin VB.Menu menu03_04_02 
            Caption         =   "Codigo Cajas"
         End
      End
      Begin VB.Menu menu03_05 
         Caption         =   "Procesos"
         Begin VB.Menu menu03_05_01 
            Caption         =   "Anulacion recibos"
         End
         Begin VB.Menu menu03_05_02 
            Caption         =   "Regenera saldos"
         End
      End
   End
   Begin VB.Menu menu04 
      Caption         =   "Configuracion"
      Begin VB.Menu menu04_01 
         Caption         =   "Usuarios"
      End
      Begin VB.Menu menu04_02 
         Caption         =   "Pto vta x usuario"
      End
   End
   Begin VB.Menu menu05 
      Caption         =   "Salir"
   End
End
Attribute VB_Name = "MDIPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adoreg As ADODB.Recordset
Dim rs As ADODB.Recordset
Private Sub Cmd3_Click()
   VGRegEnt = 1
   FrmRegistro.Show
End Sub

Private Sub Cmd4_Click()
  VGGuiaSal = True
  VGRegEnt = 2
  FrmGuiaSal.Show
End Sub

Private Sub Cmd9_Click()
If MsgBox("Esta seguro que desea salir?", vbYesNo + vbInformation, "Sistemas") = vbYes Then End

End Sub

Private Sub Form_Load()
Dim sFileName As String
Dim sBD As String
Dim sBDt As String
Dim n As String
Dim RSQL As String
Dim IASA As String
On Error GoTo Err
Set VGdllApi = New dll_apisgen.dll_apis
   
'Verificar_Sistema
VGCodMon = "01"
VGtransp = True
VGSALIR = False
VGcomputer = UCase(ComputerName)
VGsql = VGdllApi.LeerIni(App.Path & "\Marfice.ini", "conexion", "SQL", "")
VGsql = IIf(VGsql = "", 0, VGsql)
   
GPunto = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "CONEXION", "PUNTOVTA", "?")
GPunto = IIf(GPunto = "?", "01", GPunto)
g_ptoventa = GPunto

VGformatofecha = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "CONEXION", "FECHASQL", "?")
VGformatofecha = IIf(VGformatofecha = "?", "MDY", VGformatofecha)
       
'Conexion de General
VGParamSistem.BDEmpresaGEN = Trim(VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "BDGENERAL", "BDDATOS", "?"))
VGParamSistem.ServidorGEN = Trim(VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "BDGENERAL", "SERVIDOR", "?"))
VGParamSistem.UsuarioGEN = Trim(VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "BDGENERAL", "USUARIO", "?"))
' VGParamSistem.PwdGEN = DECODIFICA(Trim(VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "BDGENERAL", "PASSW", "?")), NUMMAGICO)
VGParamSistem.PwdGEN = Trim(VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "BDGENERAL", "PASSW", "?"))
        
'Conexion de inventarios
VGParamSistem.BDEmpresa = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "conexion", "BDDATOS", "?")
VGParamSistem.Servidor = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "conexion", "SERVIDOR", "?")
VGParamSistem.Usuario = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "conexion", "USUARIO", "?")
' VGParamSistem.PWD = DECODIFICA(Trim(VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "conexion", "PASSW", "?")), NUMMAGICO)
VGParamSistem.Pwd = Trim(VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "conexion", "PASSW", "?"))

VGOrden = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "conexion", "ORDEN", "?")
   
   ' reportes
VGParamSistem.RutaReport = Trim(VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "REPORTES", "PYME", "?"))
VGParamSistem.carpetareportes = Trim(VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "conexion", "CARPETAREPORTES", "?"))
   
'Conexion de Contabilidad
VGParamSistem.BDEmpresaCT = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "CONTABILIDAD", "BDDATOS", "?")
If VGParamSistem.BDEmpresaCT = "" Then
   VGParamSistem.BDEmpresaCT = VGParamSistem.BDEmpresa
   VGParamSistem.ServidorCT = VGParamSistem.Servidor
   VGParamSistem.UsuarioCT = VGParamSistem.Usuario
   VGParamSistem.PwdCT = VGParamSistem.Pwd
Else
   VGParamSistem.BDEmpresaCT = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "CONTABILIDAD", "BDDATOS", "?")
   VGParamSistem.ServidorCT = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "CONTABILIDAD", "SERVIDOR", "?")
   VGParamSistem.UsuarioCT = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "CONTABILIDAD", "USUARIO", "?")
 '   VGParamSistem.PwdCT = DECODIFICA(Trim(VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "CONTABILIDAD", "PASSW", "?")), NUMMAGICO)
   VGParamSistem.PwdCT = Trim(VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "CONTABILIDAD", "PASSW", "?"))

End If

If VGParamSistem.RutaReport = "" Or VGParamSistem.RutaReport = "?" Then
   VGParamSistem.RutaReport = App.Path
   VGParamSistem.carpetareportes = "Reportes"
End If
       
'Establecer Cadena de Conexión de Reportes
VGCadenaReport2 = "DSN=jckconsultores;DSQ=" & VGParamSistem.BDEmpresaGEN & ";UID=" & VGParamSistem.UsuarioGEN & ";PWD=" & VGParamSistem.PwdGEN & ""
          
mensaje1 = "Prueba - Inventarios"
sFileName = "marfice.ini"
VGDIRE = sGetIni("Marfice.ini", "CONFIG", "DIRE", "?")

frmlogin.Show 1
MDIPrincipal.Caption = "Sistema de Inventario Empresa : " & VGParametros.NomEmpresa & "   Base de datos --> " & VGParamSistem.BDEmpresa

If VGSALIR Then
   If VGCNx.State = 1 Then VGCNx.Close
   If VGcnxCT.State = 1 Then VGcnxCT.Close
      MDIPrincipal.Visible = False
      Unload Me
      Exit Sub
Else
      Call ParametrosdeAlmacenes
End If

VGAutomatico = False


Exit Sub

Err:
    MsgBox Err.Description, vbExclamation, "Aviso"
    Exit Sub
    Resume
End Sub



Private Sub menu04_08_Click(Index As Integer)
If VGParametros.PermiteRequerimientos Then
    FrmOrdenes_Requerimientos.Show 1
 Else
    frmOrdenes.Show 1
End If
End Sub

Private Sub menu01_01_01_Click()
   VGRegEnt = 1
   FrmRegistro.Show
End Sub

Private Sub menu01_01_02_Click()
   VGRegEnt = 0
   FrmRegistro.Show
End Sub

Private Sub menu01_01_03_Click()
FrmTraslado.Show
End Sub



Private Sub menu01_01_04_Click()
frmRequerimientosPedidos.Show
End Sub

Private Sub menu01_02_01_Click()
FrmArArticulo.Show
End Sub

Private Sub menu01_02_02_Click()
FrmAlmacen.Show
End Sub

Private Sub menu01_02_03_Click()
FrmTransaccion.Show
End Sub

Private Sub menu01_02_04_Click()
FrmCfgDocumento.Show
End Sub

Private Sub menu01_02_05_Click()
FrmMntUnidMedida.Show
End Sub

Private Sub menu01_02_06_Click()
FrmMntFamilia.Show
End Sub

Private Sub menu01_03_01_Click()
FrmLisArticulos.Show
End Sub



Private Sub menu01_03_03_Click()
FrmStockAlmacen.Show
End Sub

Private Sub menu01_03_04_Click()
FrmKardex.Show
End Sub

Private Sub menu01_03_05_01_Click()
FrmDocuDeta.Show
End Sub

Private Sub menu01_03_05_02_Click()
FrmRepDocuResumen.Show 1
End Sub

Private Sub menu01_03_05_03_Click()
FrmRepTraslados.Show
End Sub

Private Sub menu01_04_01_Click()
FrmSaldosConsolidados.Show
End Sub

Private Sub menu01_04_04_Click()
FrmConsultaNotas.Show
End Sub

Private Sub menu01_05_01_Click()
FrmPrcSaldos.Show
End Sub

Private Sub menu02_01_01_Click()
FrmPedidoGlosa.Show
End Sub

Private Sub menu02_01_02_Click()
VgModificar = 0
Frmcorrecciondatosgen.Show
End Sub

Private Sub menu02_01_03_Click()
VgModificar = 1
Frmcorrecciondatosgen.Show
End Sub

Private Sub menu02_03_01_Click()
FrmVtasporPuntoVta.resumen = 0
FrmVtasporPuntoVta.Show
End Sub

Private Sub menu02_03_02_Click()
FrmVtasporPuntoVta.resumen = 1
FrmVtasporPuntoVta.Show
End Sub

Private Sub menu02_04_01_Click()
FrmPuntoVenta.Show
End Sub

Private Sub menu02_04_02_Click()
FrmPtoVtaDoc.Show
End Sub

Private Sub menu02_04_03_Click()
FrmModoVenta.Show
End Sub

Private Sub menu02_04_04_Click()
FrmListaPrecios.Show
End Sub

Private Sub menu02_04_05_Click()
Frmcliente.Show
End Sub

Private Sub menu03_01_01_Click()
FrmPlanillaVarios.Show
End Sub

Private Sub menu03_01_02_Click()
FrmMovimientoClientes.Show
End Sub

Private Sub menu03_01_03_Click()
FrmMovimientoCaja.Show
End Sub

Private Sub menu03_01_05_Click()
FrmPlanillaVariosModi.Show
End Sub


Private Sub menu03_02_01_Click()
FrmSaldoxCliente.Show
End Sub

Private Sub menu03_02_02_Click()
FrmLiquidarionDiaria.resumido = 0
FrmLiquidarionDiaria.Show
End Sub

Private Sub menu03_02_03_Click()
FrmLiquidarionDiaria.resumido = 1
FrmLiquidarionDiaria.Show
End Sub

Private Sub menu03_02_04_Click()
FrmImprimirRecibo.Show
End Sub


Private Sub menu03_03_02_Click()
FrmBusqueda.Show
End Sub

Private Sub menu03_04_01_Click()
FrmEmpresa.Show
End Sub

Private Sub menu03_04_02_Click()
FrmCodigocajas.Show
End Sub

Private Sub menu03_05_01_Click()
frmAnularBorraRecibos.Show
End Sub

Private Sub menu03_05_02_Click()
FrmGeneraSaldos.Show
End Sub

Private Sub menu04_01_Click()
Frmusuarios.Show
End Sub

Private Sub menu04_02_Click()
Frmusuariosxpuntovta.Show
End Sub

Private Sub menu05_Click()
Unload Me
End Sub
