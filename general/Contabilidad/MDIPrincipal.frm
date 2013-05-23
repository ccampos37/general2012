VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.MDIForm MDIPrincipal 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Sistema de Contabilidad"
   ClientHeight    =   7500
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   9960
   Icon            =   "MDIPrincipal.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
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
      Top             =   7170
      Width           =   9960
      _ExtentX        =   17568
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
            Picture         =   "MDIPrincipal.frx":0CCA
            Text            =   "Tipo Cambio"
            TextSave        =   "Tipo Cambio"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4480
            MinWidth        =   4480
            Picture         =   "MDIPrincipal.frx":0FE6
            Text            =   "Servidor"
            TextSave        =   "Servidor"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7832
            MinWidth        =   7832
            Picture         =   "MDIPrincipal.frx":1142
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
            Picture         =   "MDIPrincipal.frx":129E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":16F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":184C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":19B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":1B10
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":1C6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":1DC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":1F24
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":2084
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
      Width           =   9960
      _ExtentX        =   17568
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
         Visible         =   0   'False
      End
      Begin VB.Menu mnu00_01 
         Caption         =   "Eliminar"
         Index           =   3
         Shortcut        =   ^E
         Visible         =   0   'False
      End
      Begin VB.Menu mnu00_01 
         Caption         =   "Modificar"
         Index           =   4
         Shortcut        =   ^U
         Visible         =   0   'False
      End
      Begin VB.Menu mnu00_01 
         Caption         =   "Cancelar"
         Index           =   5
         Shortcut        =   {F3}
         Visible         =   0   'False
      End
      Begin VB.Menu mnu00_01 
         Caption         =   "Insertar detalle"
         Index           =   6
         Shortcut        =   {F5}
         Visible         =   0   'False
      End
      Begin VB.Menu mnu00_01 
         Caption         =   "Eliminar detalle"
         Index           =   7
         Shortcut        =   {F6}
         Visible         =   0   'False
      End
      Begin VB.Menu mnu00_01 
         Caption         =   "Imprimir"
         Index           =   8
         Shortcut        =   ^P
         Visible         =   0   'False
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
   Begin VB.Menu mnu01 
      Caption         =   "&Tablas Básicas"
      Begin VB.Menu mnu01_01 
         Caption         =   "Principales"
         Begin VB.Menu mnu01_01_02 
            Caption         =   "&Operación"
         End
         Begin VB.Menu mnu01_01_03 
            Caption         =   "&Centro Costos"
         End
         Begin VB.Menu mnu01_01_08 
            Caption         =   "Tipo de &Documento"
         End
         Begin VB.Menu mnu01_01_09 
            Caption         =   "&Estado Comprobante"
         End
         Begin VB.Menu mnu01_01_10 
            Caption         =   "A&plicación"
         End
         Begin VB.Menu mnu01_01_14 
            Caption         =   "Tipo de Monedas"
         End
         Begin VB.Menu mnu01_01_15 
            Caption         =   "Tipo de Cuentas "
         End
      End
      Begin VB.Menu mnu01_04 
         Caption         =   "Analitico"
         Begin VB.Menu mnu01_04_01 
            Caption         =   "Tipo de Analítico"
         End
         Begin VB.Menu mnu01_04_02 
            Caption         =   "&Entidad"
         End
      End
      Begin VB.Menu mnu01_05 
         Caption         =   "Libro"
      End
      Begin VB.Menu mnu01_06 
         Caption         =   "&Asiento y Sub Asientos"
         Begin VB.Menu mnu01_06_01 
            Caption         =   "Asiento"
         End
         Begin VB.Menu mnu01_06_02 
            Caption         =   "Sub Asiento"
         End
         Begin VB.Menu mnu01_06_03 
            Caption         =   "Plantilla Asiento"
         End
      End
      Begin VB.Menu mnu01_07 
         Caption         =   "Tipo &Cambio"
      End
      Begin VB.Menu mnu01_11 
         Caption         =   "Estructuras"
         Begin VB.Menu mnu01_11_01 
            Caption         =   "Estructura del Balance"
         End
         Begin VB.Menu mnu01_11_02 
            Caption         =   "Estructura del Estado de Ganacias y Pérdidas"
         End
         Begin VB.Menu mnu01_11_03 
            Caption         =   "Totalizador Línea E.G.P"
         End
         Begin VB.Menu mnu01_11_04 
            Caption         =   "&Ratios Financieros"
            Visible         =   0   'False
         End
         Begin VB.Menu mnu01_11_05 
            Caption         =   "Parámetros Gastos"
         End
         Begin VB.Menu mnu01_11_06 
            Caption         =   "Parámetros Libro Auxiliar"
         End
      End
      Begin VB.Menu mnu01_12 
         Caption         =   "Plan de Cuentas"
      End
      Begin VB.Menu mnu01_13 
         Caption         =   "Saldos Iniciales"
      End
      Begin VB.Menu mnu01_02 
         Caption         =   "Libros Electronicos"
         Begin VB.Menu mnu01_02_01 
            Caption         =   "Oportunidad de Presentacion"
         End
         Begin VB.Menu mnu01_02_02 
            Caption         =   "Indicador de Operacion"
         End
      End
   End
   Begin VB.Menu mnu02 
      Caption         =   "&Movimientos/consultas"
      Begin VB.Menu mnu02_01 
         Caption         =   "Consulta de Comprobantes"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu02_02 
         Caption         =   "Mantenimiento de Comprobantes"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnu02_04 
         Caption         =   "Apertura Cuenta Corriente"
      End
      Begin VB.Menu mnu02_03 
         Caption         =   "Conciliación Bancaria"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnu03 
      Caption         =   "&Procesos"
      Begin VB.Menu mnu03_01 
         Caption         =   "Mayorizar Mes Activo"
      End
      Begin VB.Menu mnu03_02 
         Caption         =   "Ajuste de Diferencia de cambiox Doc"
      End
      Begin VB.Menu mnu03_09 
         Caption         =   "Generar &Saldos Iniciales"
      End
      Begin VB.Menu mnu03_04 
         Caption         =   "Importar"
         Begin VB.Menu mnu03_04_01 
            Caption         =   "&Importar Datos Facturacion"
         End
         Begin VB.Menu mnu03_04_02 
            Caption         =   "&Importar Datos Cobranza"
         End
         Begin VB.Menu mnu03_04_03 
            Caption         =   "Importar Datos Pagar"
         End
      End
      Begin VB.Menu mnu03_07 
         Caption         =   "Generar Saldos Analiticos"
      End
   End
   Begin VB.Menu mnu04 
      Caption         =   "Reportes"
      Begin VB.Menu mnu04_01 
         Caption         =   "Diario General"
      End
      Begin VB.Menu mnu04_02 
         Caption         =   "Balance de Comprobación"
      End
      Begin VB.Menu mnu04_03 
         Caption         =   "Mayor Analítico"
      End
      Begin VB.Menu mnu04_04 
         Caption         =   "Mayor General"
      End
      Begin VB.Menu mnu04_05 
         Caption         =   "Libros Auxiliares"
         Begin VB.Menu mnu04_05_01 
            Caption         =   "Registros de Compras"
         End
         Begin VB.Menu mnu04_05_02 
            Caption         =   "Registros de Ventas"
         End
         Begin VB.Menu mnu04_05_03 
            Caption         =   "Honorarios"
         End
         Begin VB.Menu mnu04_05_04 
            Caption         =   "Caja y Bancos"
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mnu04_06 
         Caption         =   "Saldos Analíticos"
      End
      Begin VB.Menu mnu04_07 
         Caption         =   "Varios"
         Begin VB.Menu mnu04_07_01 
            Caption         =   "Consistencias"
            Begin VB.Menu mnu04_07_01_01 
               Caption         =   "Saldos Iniciales"
            End
            Begin VB.Menu mnu04_07_01_02 
               Caption         =   "Consistencia de Asientos"
            End
            Begin VB.Menu mnu04_07_01_03 
               Caption         =   "Consistencia de Ventas"
               Visible         =   0   'False
            End
            Begin VB.Menu mnu04_07_01_04 
               Caption         =   "Diferencias Cta. Ctble. vs. Cta. Análisis"
            End
         End
         Begin VB.Menu mnu04_07_02 
            Caption         =   "Tablas"
            Begin VB.Menu mnu04_07_02_01 
               Caption         =   "Plantilla de Sub Asientos"
            End
            Begin VB.Menu mnu04_07_02_02 
               Caption         =   "Cuentas Distribución"
            End
            Begin VB.Menu mnu04_07_02_03 
               Caption         =   "Plan de Cuentas"
               Visible         =   0   'False
            End
            Begin VB.Menu mnu04_07_02_04 
               Caption         =   "Estructuras"
            End
         End
         Begin VB.Menu mnu04_07_04 
            Caption         =   "Comprobantes por rangos"
         End
         Begin VB.Menu mnu04_07_07 
            Caption         =   "Estados Financieros"
         End
         Begin VB.Menu mnu04_07_09 
            Caption         =   "Movimientos Cuentas"
         End
      End
      Begin VB.Menu mnu04_08 
         Caption         =   "Centro de Costos"
         Begin VB.Menu mnu04_08_01 
            Caption         =   "Mensual"
         End
         Begin VB.Menu mnu04_08_02 
            Caption         =   "Acumulados"
         End
      End
      Begin VB.Menu mnu04_09 
         Caption         =   "Inventarios y Nalances"
      End
      Begin VB.Menu mnu04_10 
         Caption         =   "Nuevos Libros Tributarios"
         Begin VB.Menu mnu04_10_02 
            Caption         =   "Balance de Comprobacion"
         End
         Begin VB.Menu mnu04_10_03 
            Caption         =   "Libro de inventarios y Balances"
         End
         Begin VB.Menu mnu04_10_05 
            Caption         =   "Libro Diario"
         End
         Begin VB.Menu mnu04_10_06 
            Caption         =   "Libro Mayor"
         End
         Begin VB.Menu mnu04_10_07 
            Caption         =   "Libro Caja y Bancos"
         End
         Begin VB.Menu mnu04_10_08 
            Caption         =   "Registro de Compras"
         End
         Begin VB.Menu mnu04_10_10 
            Caption         =   "Estados Financieros"
         End
         Begin VB.Menu mnu04_10_14 
            Caption         =   "Regisro de Ventas e Ingresos"
         End
      End
   End
   Begin VB.Menu mnu05 
      Caption         =   "&Configuración"
      Begin VB.Menu mnu05_01 
         Caption         =   "Aperturar Año"
      End
      Begin VB.Menu mnu05_02 
         Caption         =   "&Parámetros Generales"
      End
      Begin VB.Menu mnu05_03 
         Caption         =   "C&reacion de Usuarios"
      End
      Begin VB.Menu mnu05_04 
         Caption         =   "&Configuracion de Empresas"
         Begin VB.Menu mnu05_04_01 
            Caption         =   "Creacion de empresas"
         End
      End
      Begin VB.Menu mnu05_05 
         Caption         =   "Cierres mensuales"
      End
   End
   Begin VB.Menu mnu06 
      Caption         =   "Sunat"
      Begin VB.Menu mnu06_01 
         Caption         =   "Formularios"
         Begin VB.Menu mnu06_01_01 
            Caption         =   "Form 682"
         End
      End
      Begin VB.Menu mnu06_02 
         Caption         =   "Libros Electronicos"
         Begin VB.Menu mnu06_02_01 
            Caption         =   "Libros Principales"
         End
         Begin VB.Menu mnu06_02_02 
            Caption         =   "Libros auxiliares"
         End
         Begin VB.Menu mnu06_02_03 
            Caption         =   "Librios Inventarios y Balamcer"
         End
         Begin VB.Menu mnu06_02_14 
            Caption         =   "Estados Financieros"
         End
         Begin VB.Menu mnu06_02_10 
            Caption         =   "Otros"
         End
      End
   End
   Begin VB.Menu mnu08 
      Caption         =   "&Salir"
      WindowList      =   -1  'True
   End
End
Attribute VB_Name = "MDIPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub MDIForm_Load()
      
Call ADOConectar
Call ADOConectarReport("CONTABILIDAD")
VGtipo = contab

mensaje1 = "Prueba "

frmlogin.Show 1
MDIPrincipal.Caption = "Sistema de Contabilidad Empresa : " & VGParametros.NomEmpresa & "   Base de datos --> " & VGParamSistem.BDEmpresa

If VGSalir Then
   If VGCNx.State = 1 Then VGCNx.Close
   If VGCnxCT.State = 1 Then VGCnxCT.Close
      MDIPrincipal.Visible = False
      Unload Me
      Exit Sub
Else
      Call CargarParametrosContabilidad
End If

Exit Sub

err:
    MsgBox err.Description, vbExclamation, "Aviso"
    Exit Sub
    Resume
Xmain:
    MsgBox err.Description, vbExclamation

End Sub

Private Sub mnu00_01_01_Click(Index As Integer)
'FIXIT: Siempre que sea posible, reemplace ActiveForm o ActiveControl con una variable en tiempo de compilación.     FixIT90210ae-R1614-RCFE85
    Call Screen.ActiveForm.Pavant(Index)
End Sub

Private Sub mnu00_01_Click(Index As Integer)
'FIXIT: Siempre que sea posible, reemplace ActiveForm o ActiveControl con una variable en tiempo de compilación.     FixIT90210ae-R1614-RCFE85
    Call Screen.ActiveForm.PMant(Index)
End Sub

Private Sub mnu01_01_02_Click()
    frmMantOperacion.Show
End Sub

Private Sub mnu01_01_03_Click()
    frmMantCentroCosto.Show
End Sub


Private Sub mnu01_02_02_Click()
FrmMntIndicadorOportunidad.Show
End Sub

Private Sub mnu01_04_01_Click()
    frmMantTipoAnalitico.Show
End Sub

Private Sub mnu01_04_02_Click()
    frmMantEntidad.Show
End Sub

Private Sub mnu01_01_05_Click()
    frmMantLibro.Show
End Sub

Private Sub mnu01_06_01_Click()
    frmMantAsiento.Show
End Sub

Private Sub mnu01_06_02_Click()
 If ValidaAsientos = True Then
   frmMantSubAsiento.Show
 End If
End Sub

Private Sub mnu01_06_03_Click()
 If ValidaAsientos = True Then
    If ValidaSubAsientos("%") = True Then frmMantPlantillaAsiento.Show
 End If
End Sub

Private Sub mnu01_07_Click()
'    frmPassword.Show
    frmMantTipoCambio.Show
End Sub

Private Sub mnu01_01_08_Click()
    frmMantTipoDocumento.Show
End Sub

Private Sub mnu01_01_09_Click()
    frmMantEstComprobante.Show
End Sub

Private Sub mnu01_01_10_Click()
    frmMantAplicacion.Show
End Sub

Private Sub mnu01_11_01_Click()
    frmEstructuraMantBalance.Show
End Sub

Private Sub mnu01_11_02_Click()
    frmEstructuraMantEstadoGanPer.Show
End Sub

Private Sub mnu01_11_03_Click()
    frmEstructuraMantTotalLineaEGP.Show
End Sub

Private Sub mnu01_11_05_Click()
    frmEstructuraMantParametrosGastos.Show
End Sub

Private Sub mnu01_11_06_Click()
    frmEstructuraMantParamLibAux.Show
End Sub

Private Sub mnu01_12_Click()
   frmMantPlanCuentas.Show
End Sub

Private Sub mnu01_13_Click()
    frmMantSaldosInicialPlan.Show
End Sub

Private Sub mnu01_01_14_Click()
    FrmMantMoneda.Show
End Sub

Private Sub mnu01_01_15_Click()
FrmTipocuenta.Show 1
End Sub

Private Sub mnu02_01_Click()
    frmConsultaComprobantes.Show
End Sub

Private Sub mnu02_02_Click()
    Screen.MousePointer = vbHourglass
    frmantcomprobantes.Show
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnu02_03_Click()
'     FrmConciliacion.Show
End Sub

Private Sub mnu02_04_Click()
   frmMant_CtaCteAnalitico.Show 1
End Sub

Private Sub mnu03_01_Click()
  On Error GoTo Mayor
    Screen.MousePointer = 11
    VGCNx.BeginTrans
    Set VGCommandoSP = New ADODB.Command
    VGCommandoSP.ActiveConnection = VGGeneral
    VGCommandoSP.CommandType = adCmdStoredProc
    VGCommandoSP.CommandText = "ct_mayoriza_pro"
    VGCommandoSP.Parameters.Refresh
    With VGCommandoSP
        .Parameters("@base") = VGParamSistem.BDEmpresa
        .Parameters("@empresa") = VGParametros.empresacodigo
        .Parameters("@anno") = VGParamSistem.Anoproceso
        .Parameters("@mespro") = VGParamSistem.Mesproceso
        .Parameters("@user") = VGParamSistem.Usuario
        .Execute
    End With
    VGCNx.CommitTrans
    Screen.MousePointer = 1
    MsgBox "Se Mayorizo Satisfactoriamente ", vbInformation
    Exit Sub
Mayor:
    Screen.MousePointer = 1
    VGCNx.RollbackTrans
    MsgBox "No se pudo mayorizar " & Chr(13) & err.Description, vbExclamation
End Sub
Private Sub mnu03_03_Click()
    Call CancelaDocumentos
End Sub
Private Sub mnu03_02_Click()
FrmAjusDiferxDoc.Show
End Sub

Private Sub mnu03_04_01_Click()
    FrmImportDataFac.Show
End Sub

Private Sub mnu03_04_02_Click()
    FrmContabCobrar.Show
End Sub

Private Sub mnu03_07_Click()
FrmgenerasaldosAnaliticos.Show
End Sub

Private Sub mnu03_04_03_Click()
FrmContabPagar.Show 1
End Sub

Private Sub mnu03_09_Click()
FrmGenerasaldosini.Show 1
End Sub

Private Sub mnu04_01_Click()
  frmRepDiarioGeneral.Show
End Sub

Private Sub mnu04_02_Click()
  FrmRepBalanceComp.Show
End Sub

Private Sub mnu04_03_Click()
  frmRepMayor.Caso = "1"
  frmRepMayor.tituloreporte = "Reporte de Mayor Analítico"
  frmRepMayor.Show
End Sub

Private Sub mnu04_04_Click()
  frmRepMayor.Caso = "2"
  frmRepMayor.tituloreporte = "Reporte de Mayor General"
  frmRepMayor.Show
End Sub

Private Sub mnu04_05_01_Click()
    FrmRepCompras.Show
End Sub

Private Sub mnu04_05_02_Click()
  frmRepVentas.Show
End Sub

Private Sub mnu04_05_03_Click()
  frmRepHonorarios.Show
End Sub

Private Sub mnu04_05_04_Click()
    frmRepCajaBancos.Show
End Sub

Private Sub mnu04_06_Click()
    frmRepCtaCteAnalitico.Show
End Sub

Private Sub mnu04_07_01_01_Click()
     frmRepPlanCuentasSaldosIniciales.Show
End Sub

Private Sub mnu04_07_01_02_Click()
 Dim SQL As String
  On Error GoTo xx
    
'FIXIT: Declare 'arrform' con un tipo de datos de enlace en tiempo de compilación          FixIT90210ae-R1672-R1B8ZE
  Dim arrform(0) As Variant, arrparm(4) As Variant
  Dim NombreRep As String, CadOrden As String
    arrparm(0) = VGParamSistem.BDEmpresa
    arrparm(1) = VGParametros.empresacodigo
    arrparm(2) = VGParamSistem.Anoproceso
    arrparm(3) = Trim$(VGParamSistem.Mesproceso)
    NombreRep = "rptAsientosDescuadrados.rpt"
    Call ImpresionRptProc(NombreRep, arrform, arrparm, CadOrden, "Descuadre de Asientos")
    Exit Sub
xx:
    MsgBox "No se pudo Abrir el Reporte " & Chr(13) & err.Description, vbExclamation
End Sub


Private Sub mnu04_07_01_04_Click()
    FrmRepCuentasVsAnaliticos.Show
End Sub

Private Sub mnu04_07_02_01_Click()
  frmRepPlantillaSubAsientos.Show
End Sub

Private Sub mnu04_07_02_02_Click()
 frmRepListadoCtasDist.Show
End Sub


Private Sub mnu04_07_03_Click()
  frmRepPlanCuentas.Show
End Sub

Private Sub mnu04_07_02_04_Click()
  frmRepEstructuras.Show
End Sub

Private Sub mnu04_07_04_Click()
    frmRepComprob.Show
End Sub


Private Sub mnu04_07_07_Click()
  frmRepEstadosFinancieros.Show
End Sub


Private Sub mnu04_07_09_Click()
    frmRepMovimientoCuentas.Show
End Sub

Private Sub mnu04_07_10_Click()
    frmValidacionSeries.Show
End Sub

Private Sub mnu04_08_01_Click()
  frmRepCentrodeCostos.Show
End Sub

Private Sub mnu04_08_02_Click()
frmRepCentrodeCostosAcumulado.Show 1
End Sub
Private Sub mnu04_09_Click()
FrmInventariosyBalances.Show
End Sub


Private Sub mnu04_10_02_Click()
FrmLibroBalancedeComprobacion.Show
End Sub

Private Sub mnu04_10_03_Click()
FrmLibroInventariosyBalances.Show
End Sub

Private Sub mnu04_10_05_Click()
FrmLibroDiario.Show
End Sub

Private Sub mnu04_10_06_Click()
FrmLibroMayor.Caso = "1"
FrmLibroMayor.tituloreporte = "Reporte de LIBRO MAYOR"
FrmLibroMayor.Show
End Sub

Private Sub mnu04_10_07_Click()
FrmLibroCajayBancos.Show
End Sub

Private Sub mnu04_10_08_Click()
FrmLibroRegistrodeCompras.Show
End Sub

Private Sub mnu04_10_10_Click()
frmRepEstadosFinancieros.Show
End Sub

Private Sub mnu04_10_14_Click()
FrmLibroRegistrodeventas.Show
End Sub

Private Sub mnu04_10_15_Click()
FrmLibroInventarios.Show
End Sub
Private Sub mnu05_01_Click()
    frmannos.Show
End Sub
Private Sub mnu05_02_Click()
    frmParametros.Show
End Sub
Private Sub mnu05_03_Click()
   Frmusuarios.Show
End Sub

Private Sub mnu05_04_01_Click()
FrmCreacionEmpresa.Show
End Sub

Private Sub mnu05_04_Click()
VGtipo = contab
'FrmCfgEmpresa.Show
End Sub

Private Sub mnu05_05_Click()
FrmCierremensual.Show 1
End Sub

Private Sub mnu06_01_01_Click()
FrmSunat682.Show
End Sub

Private Sub mnu06_02_01_Click()
FrmLibrosElectPrincipales.Show
End Sub

Private Sub mnu08_Click()
Unload Me
End Sub

Private Sub StatusBar1_PanelDblClick(ByVal Panel As MSComctlLib.Panel)
    Select Case Panel.Index
        Case 1, 2
            frmselanomes.Show 1
    End Select
End Sub
Private Sub ToolComprob_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case CInt(Right(Trim$(Button.Key), Len(Trim$(Button.Key)) - 1))
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
