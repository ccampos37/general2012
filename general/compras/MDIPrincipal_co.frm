VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.MDIForm MDIPrincipal 
   BackColor       =   &H8000000C&
   Caption         =   "Sistema de Provisión de Compras"
   ClientHeight    =   8205
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   10215
   Icon            =   "MDIPrincipal_co.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   144
      Top             =   6576
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport cryRpt 
      Left            =   8700
      Top             =   7365
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   7875
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
            Text            =   "Mes Proceso"
            TextSave        =   "Mes Proceso"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
            Text            =   "Año Proceso"
            TextSave        =   "Año Proceso"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4657
            MinWidth        =   4657
            Text            =   "Fecha de Trabajo"
            TextSave        =   "Fecha de Trabajo"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4304
            MinWidth        =   4304
            Picture         =   "MDIPrincipal_co.frx":1272
            Text            =   "Tipo Cambio"
            TextSave        =   "Tipo Cambio"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4480
            MinWidth        =   4480
            Picture         =   "MDIPrincipal_co.frx":158E
            Text            =   "Servidor"
            TextSave        =   "Servidor"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7832
            MinWidth        =   7832
            Picture         =   "MDIPrincipal_co.frx":16EA
            Text            =   "Base de Datos"
            TextSave        =   "Base de Datos"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9180
      Top             =   7215
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
            Picture         =   "MDIPrincipal_co.frx":1846
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal_co.frx":1C98
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal_co.frx":1DF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal_co.frx":1F5C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal_co.frx":20B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal_co.frx":2214
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal_co.frx":2370
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal_co.frx":24CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal_co.frx":262C
            Key             =   ""
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
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "k1"
            Description     =   "Nuevo"
            Object.ToolTipText     =   "Nuevo"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "k2"
            Description     =   "Grabar Salir"
            Object.ToolTipText     =   "Grabar y Salir"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "k3"
            Description     =   "Eliminar"
            Object.ToolTipText     =   "Eliminar"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "k4"
            Description     =   "Modificar"
            Object.ToolTipText     =   "Modificar"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "k5"
            Description     =   "Cancelar"
            Object.ToolTipText     =   "Cancelar Operacion"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "k6"
            Description     =   "Añadir Detalle"
            Object.ToolTipText     =   "Añadir Detalle"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "k7"
            Description     =   "Eliminar Detalle"
            Object.ToolTipText     =   "Eliminar Detalle"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "k8"
            Description     =   "Imprimir"
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin Crystal.CrystalReport CryRptProc 
      Left            =   9750
      Top             =   7350
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Menu mnu00 
      Caption         =   "&Edicion"
      Visible         =   0   'False
      Begin VB.Menu mnu00_01 
         Caption         =   "Nuevo"
         Index           =   1
         Shortcut        =   ^N
      End
      Begin VB.Menu mnu00_01 
         Caption         =   "Grabar"
         Index           =   2
         Shortcut        =   ^G
      End
      Begin VB.Menu mnu00_01 
         Caption         =   "Eliminar"
         Index           =   3
         Shortcut        =   ^E
      End
      Begin VB.Menu mnu00_01 
         Caption         =   "Modificar"
         Index           =   4
         Shortcut        =   ^U
      End
      Begin VB.Menu mnu00_01 
         Caption         =   "Cancelar"
         Index           =   5
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnu00_01 
         Caption         =   "Insertar detalle"
         Index           =   6
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnu00_01 
         Caption         =   "Eliminar detalle"
         Index           =   7
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnu00_01 
         Caption         =   "Imprimir"
         Index           =   8
         Shortcut        =   ^P
      End
      Begin VB.Menu mnu00_01 
         Caption         =   "Avanzados"
         Index           =   9
         Visible         =   0   'False
         Begin VB.Menu mnu00_01_01 
            Caption         =   "Ir al monto"
            Index           =   1
            Shortcut        =   {F9}
         End
         Begin VB.Menu mnu00_01_01 
            Caption         =   "Ir a la operacion"
            Index           =   2
            Shortcut        =   {F8}
         End
      End
   End
   Begin VB.Menu menu01 
      Caption         =   "&Tablas Básicas"
      Begin VB.Menu menu01_01 
         Caption         =   "Proveedores "
      End
      Begin VB.Menu menu01_02 
         Caption         =   "Mantenimento de Provisiones"
      End
      Begin VB.Menu menu01_03 
         Caption         =   "Mantenimiento Plan de Gastos"
      End
      Begin VB.Menu menu01_04 
         Caption         =   "Tipo de Cambio"
      End
   End
   Begin VB.Menu menu02 
      Caption         =   "&Movimientos"
      Begin VB.Menu menu_02_01 
         Caption         =   "Ordenes de Compra"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu menu02_02 
         Caption         =   "Mantenimiento de Provision"
         Shortcut        =   ^M
      End
      Begin VB.Menu menu02_03 
         Caption         =   "Consulta de Provision"
      End
      Begin VB.Menu menu03_05 
         Caption         =   "Ordenes de Compra(1)"
      End
   End
   Begin VB.Menu menu03 
      Caption         =   "&Consultas"
      Begin VB.Menu menu03_01 
         Caption         =   "Comprobante"
      End
   End
   Begin VB.Menu menu04 
      Caption         =   "Reportes"
      Begin VB.Menu menu04_01 
         Caption         =   "&Registro de Compras"
      End
      Begin VB.Menu menu04_02 
         Caption         =   "&Listado de Compras x Cuenta"
      End
      Begin VB.Menu menu04_03 
         Caption         =   "Reporte de Prueba"
         Visible         =   0   'False
      End
      Begin VB.Menu menu04_04 
         Caption         =   "Or&denes de Compra"
      End
      Begin VB.Menu menu04_05 
         Caption         =   "Estado de Orden de Compra"
         Visible         =   0   'False
      End
      Begin VB.Menu menu04_06 
         Caption         =   "Diferencias Contabilización"
      End
      Begin VB.Menu menu04_07 
         Caption         =   "Listado de Compras x &Gastos"
      End
   End
   Begin VB.Menu menu05 
      Caption         =   "&Procesos"
      Begin VB.Menu menu05_01 
         Caption         =   "&Generar Asientos a Contabilidad"
      End
   End
   Begin VB.Menu mnu06 
      Caption         =   "&Configuración"
      Begin VB.Menu mnu06_01 
         Caption         =   "Aperturar Año"
      End
      Begin VB.Menu mnu06_02 
         Caption         =   "&Parámetros Generales"
      End
      Begin VB.Menu mnu06_03 
         Caption         =   "Creacion de Usuarios"
      End
   End
   Begin VB.Menu menu07 
      Caption         =   "&Ventanas"
      WindowList      =   -1  'True
   End
End
Attribute VB_Name = "MDIPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub MDIForm_Activate()
If VgActivalogin > 1 Then Exit Sub
     Set VGvardllgen = New dllgeneral.dll_general
 frmlogin.Show 1
 StatusBar1.Panels(1).Text = "Mes Proceso: " & VGvardllgen.DESMES(VGParamSistem.Mesproceso)      'Mesproceso
 StatusBar1.Panels(2).Text = "Año Proceso: " & VGParamSistem.Anoproceso                          'AnnoProceso
 VgActivalogin = 2
End Sub

Private Sub menu03_01_Click()
    frmConsultaComprobante.Show
End Sub

Private Sub menu03_05_Click()
    frmEmisionOC.Show 1
End Sub

Private Sub mnu_02_01_Click()
    'FrmOrdenCompra.Show
End Sub
Private Sub menu04_07_Click()
FrmRepListGastos.Show 1
End Sub
Private Sub menu05_01_Click()
    FrmGenAsiento.Show
End Sub

Private Sub mnu00_01_01_Click(Index As Integer)
    Call Screen.ActiveForm.Pavant(Index)
End Sub
Private Sub mnu00_01_Click(Index As Integer)
    Call Screen.ActiveForm.PMant(Index)
End Sub
Private Sub menu01_01_Click()
    Frmcliente.Show
End Sub

Private Sub menu01_02_Click()
    FrmCOmantprovi.Show 1
End Sub

Private Sub menu01_03_Click()
frmMantPlangastos.Show
End Sub

Private Sub menu02_02_Click()
    Screen.MousePointer = vbHourglass
    frmMantprovision.Show
    Screen.MousePointer = vbDefault
End Sub

Private Sub menu04_01_Click()
    FrmRepCOCompras.Show
End Sub

Private Sub menu04_02_Click()
    FrmRepListCuenta.Show
End Sub

Private Sub menu04_04_Click()
    FrmRepOrdCompra.Show 1
End Sub

Private Sub menu04_05_Click()
    FrmRepListOrdenCompra.Show 1
End Sub

Private Sub menu04_06_Click()
   frmRepListadoDiferenciasCompras.Show 1
End Sub

Private Sub mnu06_01_Click()
    frmannos.Show
End Sub


Private Sub mnu06_03_Click()
   frmCrearUsuarios.Show
End Sub

Private Sub mnu06_02_Click()
    frmParametros.Show
End Sub

Private Sub StatusBar1_PanelDblClick(ByVal Panel As MSComctlLib.Panel)
    Select Case Panel.Index
        Case 1, 2
            frmselanomes.Show 1
    End Select
End Sub
Private Sub ToolComprob_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case CInt(Right(Trim(Button.Key), Len(Trim(Button.Key)) - 1))
        Case 1 'Nuevo
            Call mnu00_01_Click(1)
        Case 2 'grabar
            Call mnu00_01_Click(2)
        Case 3 'Eliminar
            Call mnu00_01_Click(3)
        Case 4 'Modificar
            Call mnu00_01_Click(4)
        Case 5
            Call mnu00_01_Click(5)
        Case 6
            Call mnu00_01_Click(6)
        Case 7
            Call mnu00_01_Click(7)
        Case 8
            Call mnu00_01_Click(8)
    End Select
End Sub

