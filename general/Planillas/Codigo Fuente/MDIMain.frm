VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.MDIForm MDIPrincipal 
   BackColor       =   &H8000000C&
   Caption         =   "Planillas"
   ClientHeight    =   7665
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   11250
   Icon            =   "MDIMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CryRptProc 
      Left            =   120
      Top             =   2535
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport RpPlanCab 
      Left            =   120
      Top             =   1560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11250
      _ExtentX        =   19844
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   25
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Nuevo"
            Object.ToolTipText     =   "Nuevo Registro de ..."
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Editar"
            Object.ToolTipText     =   "Editar el registro"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Eliminar"
            Object.ToolTipText     =   "Eliminar el registro"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Imprimir"
            Object.ToolTipText     =   "Imprimir reporte"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Preliminar"
            Object.ToolTipText     =   "Vista Preliminar del Reporte"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Buscar"
            Object.ToolTipText     =   "Buscar dato"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Filtrar"
            Object.ToolTipText     =   "Filtrar datos"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copiar"
            Object.ToolTipText     =   "Copiar"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cortar"
            Object.ToolTipText     =   "Cortar"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Pegar"
            Object.ToolTipText     =   "Pegar"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Trabajadores"
            Object.ToolTipText     =   "Panel de Trabajadores"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "CCostos"
            Object.ToolTipText     =   "Panel de Centros de Costos"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Activar"
            Object.ToolTipText     =   "Mes activo"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Formatos"
            Object.ToolTipText     =   "Panel de Formatos de Boletas"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Conceptos"
            Object.ToolTipText     =   "Panel de Conceptos de Pago"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "InputBox"
            Object.ToolTipText     =   "Input de Boletas de Pago"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Adelanto"
            Object.ToolTipText     =   "Input de Adelantos de Pago"
            ImageIndex      =   17
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Planillas"
            Object.ToolTipText     =   "Panel de Planillas de Remuneraciones"
            ImageIndex      =   18
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Calculadora"
            Object.ToolTipText     =   "Calculadora de Pantalla"
            ImageIndex      =   19
         EndProperty
         BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Setup"
            Object.ToolTipText     =   "Valores de Inicio"
            ImageIndex      =   20
         EndProperty
         BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Salir"
            Object.ToolTipText     =   "Salir del Sistema"
            ImageIndex      =   25
         EndProperty
         BeginProperty Button25 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "OTRO"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "IMPOPROD"
                  Text            =   "Importar Movimientos"
               EndProperty
            EndProperty
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   45
      Top             =   870
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   27
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":0360
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":0E2C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":1180
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":14D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":1928
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":1C7C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":1FD0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":2324
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":2678
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":29CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":32A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":35FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":3A50
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":432C
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":4680
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":49A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":4CC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":511C
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":5438
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":5754
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":5BA8
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":5FFC
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":6E50
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":7A54
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":7D70
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":808C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar BarraEstado 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   7290
      Width           =   11250
      _ExtentX        =   19844
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   3598
            MinWidth        =   3598
            Picture         =   "MDIMain.frx":8968
            Text            =   "Usuario: Nando"
            TextSave        =   "Usuario: Nando"
            Key             =   "User"
            Object.ToolTipText     =   "Usuario del sistema"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Key             =   "PanelActivo"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Picture         =   "MDIMain.frx":97BC
            Text            =   "3.50"
            TextSave        =   "3.50"
            Key             =   "Dolar"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   3678
            Picture         =   "MDIMain.frx":9B10
            Text            =   "Base de datos Activa"
            TextSave        =   "Base de datos Activa"
            Key             =   "BaseDatos"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Key             =   "Periodo"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "17/06/2009"
         EndProperty
      EndProperty
   End
   Begin Crystal.CrystalReport CR1 
      Left            =   120
      Top             =   2040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Align           =   2  'Align Bottom
      Height          =   225
      Left            =   0
      TabIndex        =   2
      Top             =   7065
      Width           =   11250
      _ExtentX        =   19844
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Menu men01 
      Caption         =   "&Base de Datos"
      Begin VB.Menu men01_02 
         Caption         =   "&Derechohabientes"
      End
      Begin VB.Menu men01_03 
         Caption         =   "&Areas de Trabajo"
      End
      Begin VB.Menu men01_04 
         Caption         =   "&Centros de Costos"
      End
      Begin VB.Menu men01_01 
         Caption         =   "&Trabajadores"
         Shortcut        =   ^T
      End
      Begin VB.Menu men01_05 
         Caption         =   "&Fondos de Pensiones"
      End
      Begin VB.Menu men01_06 
         Caption         =   "&Conceptos de Remuneraciones"
      End
      Begin VB.Menu men01_07 
         Caption         =   "Cuentas Corrientes"
      End
      Begin VB.Menu men01_08 
         Caption         =   "Cronograma de Pagos"
      End
      Begin VB.Menu men01_10 
         Caption         =   "Asientos Contables"
      End
      Begin VB.Menu men01_11 
         Caption         =   "Configuracion de Adelantos"
         Visible         =   0   'False
      End
      Begin VB.Menu men01_12 
         Caption         =   "Otros Archivos"
         Begin VB.Menu men01_12_01 
            Caption         =   "Centros de &Alto Riesgo"
         End
         Begin VB.Menu men01_12_02 
            Caption         =   "&Variables"
         End
         Begin VB.Menu menu01_01_03 
            Caption         =   "B&illetes"
         End
         Begin VB.Menu men01_12_03 
            Caption         =   "&Bancos"
         End
         Begin VB.Menu men01_12_04 
            Caption         =   "&Documentos"
         End
         Begin VB.Menu men01_12_05 
            Caption         =   "&Columnas de Planillas"
         End
         Begin VB.Menu men01_12_06 
            Caption         =   "&Tipos de Trabajadores"
         End
         Begin VB.Menu men01_12_07 
            Caption         =   "&Otros Datos Informativos"
         End
         Begin VB.Menu men01_12_08 
            Caption         =   "&Centros de Tareo"
         End
         Begin VB.Menu men01_12_09 
            Caption         =   "Formulas de &Vacaciones"
         End
         Begin VB.Menu men01_12_10 
            Caption         =   "Formulas de &Gratificaciones"
         End
         Begin VB.Menu men01_12_11 
            Caption         =   "Formulas de &Cts"
         End
         Begin VB.Menu men01_12_12 
            Caption         =   "Formulas de &Utilidades"
         End
         Begin VB.Menu men01_12_13 
            Caption         =   "Categoria Evento"
         End
         Begin VB.Menu men01_12_14 
            Caption         =   "Sub Categoria Evento"
         End
         Begin VB.Menu men01_12_15 
            Caption         =   "Tipo Estudio"
         End
         Begin VB.Menu men01_12_16 
            Caption         =   "Estudios"
         End
      End
      Begin VB.Menu men01_13 
         Caption         =   "-"
      End
      Begin VB.Menu men01_14 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu Men02 
      Caption         =   "&Edición"
      Begin VB.Menu Men02_01 
         Caption         =   "&Copiar"
      End
      Begin VB.Menu Men02_02 
         Caption         =   "Cor&tar"
      End
      Begin VB.Menu Men02_03 
         Caption         =   "&Pegar"
      End
      Begin VB.Menu Men02_04 
         Caption         =   "-"
      End
      Begin VB.Menu Men02_05 
         Caption         =   "&Editar Registro"
         Shortcut        =   ^I
      End
      Begin VB.Menu Men02_06 
         Caption         =   "&Agregar registro"
         Shortcut        =   ^N
      End
      Begin VB.Menu Men02_07 
         Caption         =   "Eli&minar Registro"
         Shortcut        =   ^K
      End
      Begin VB.Menu Men02_08 
         Caption         =   "&Buscar Registro"
         Shortcut        =   ^B
      End
   End
   Begin VB.Menu Men03 
      Caption         =   "&Procesos"
      Begin VB.Menu Men03_01 
         Caption         =   "&Apertura de Mes"
      End
      Begin VB.Menu Men03_02 
         Caption         =   "&Asistencia"
         Begin VB.Menu Men03_02_01 
            Caption         =   "&Registrar Asistencia"
         End
         Begin VB.Menu Men03_02_02 
            Caption         =   "&Asistencia por Centro de Costo"
            Visible         =   0   'False
         End
         Begin VB.Menu Men03_02_03 
            Caption         =   "&Trabajadores en Varios Centros de Costo"
         End
         Begin VB.Menu Men03_02_04 
            Caption         =   "-"
         End
         Begin VB.Menu Men03_02_05 
            Caption         =   "Carga de Asistencia desde medio externo Tipo 1"
            Visible         =   0   'False
         End
         Begin VB.Menu Men03_02_06 
            Caption         =   "Carga de Asistencia desde medio externo Tipo 2"
            Visible         =   0   'False
         End
         Begin VB.Menu Men03_02_07 
            Caption         =   "Configuración de Reloj Marcador (Software Externo)"
            Visible         =   0   'False
         End
      End
      Begin VB.Menu Men03_03 
         Caption         =   "Ingreso de &Movimientos"
      End
      Begin VB.Menu Men03_04 
         Caption         =   "Adelanto de Remuneraciones"
      End
      Begin VB.Menu Men03_05 
         Caption         =   "&Adelantos Detallado"
         Visible         =   0   'False
      End
      Begin VB.Menu Men03_06 
         Caption         =   "&Calculo de Planilla de Remuneraciones"
      End
      Begin VB.Menu Men03_07 
         Caption         =   "&Renta de 5ta. Categoria"
         Begin VB.Menu Men03_07_01 
            Caption         =   "&Cálculo de Renta de 5ta. Categoria"
            Visible         =   0   'False
         End
         Begin VB.Menu Men03_07_02 
            Caption         =   "&Parámetros de Cálculo 5ta. Categoria"
         End
         Begin VB.Menu Men03_07_03 
            Caption         =   "&Trabajadores afectos a 5ta. Categoria"
         End
      End
      Begin VB.Menu Men03_08 
         Caption         =   "&Vacaciones"
      End
      Begin VB.Menu Men03_09 
         Caption         =   "&Gratificaciones"
      End
      Begin VB.Menu Men03_10 
         Caption         =   "&Liquidación de Trabajadores"
      End
      Begin VB.Menu Men03_11 
         Caption         =   "Depósitos de C.T.S."
      End
      Begin VB.Menu Men03_12 
         Caption         =   "Verificador de Consistencia"
      End
      Begin VB.Menu Men03_13 
         Caption         =   "Calculo de &Utilidades"
      End
      Begin VB.Menu Men03_14 
         Caption         =   "Provisiones"
      End
      Begin VB.Menu Men03_15 
         Caption         =   "Cierre de Mes"
      End
      Begin VB.Menu Men03_16 
         Caption         =   "Reapertura de Mes"
      End
   End
   Begin VB.Menu Men04 
      Caption         =   "Pla&nillas"
      Begin VB.Menu Men04_01 
         Caption         =   "&Panel de Planillas Mensuales procesadas"
      End
      Begin VB.Menu Men04_02 
         Caption         =   "&Enviar Planillas al Almacen"
      End
      Begin VB.Menu Men04_03 
         Caption         =   "-"
      End
      Begin VB.Menu Men04_04 
         Caption         =   "&Panel de Administración de Planillas"
      End
      Begin VB.Menu Men04_05 
         Caption         =   "Panel de Administración de Adelantos"
      End
      Begin VB.Menu Men04_06 
         Caption         =   "-"
      End
      Begin VB.Menu Men04_07 
         Caption         =   "&Formatos de Planillas"
      End
      Begin VB.Menu Men04_08 
         Caption         =   "&Imprimir Cabeceras de planiila"
      End
   End
   Begin VB.Menu Men05 
      Caption         =   "Reportes"
      Begin VB.Menu Men05_01 
         Caption         =   "Adelantos pendientes de descontar"
      End
      Begin VB.Menu Men05_02 
         Caption         =   "&Asistencia"
         Begin VB.Menu Men05_02_01 
            Caption         =   "&Informe de Asistencia por Trabajador"
         End
         Begin VB.Menu Men05_02_02 
            Caption         =   "&Informe de Asistencia Por Fechas"
         End
         Begin VB.Menu Men05_02_03 
            Caption         =   "&Consolidado para Planillas"
         End
      End
      Begin VB.Menu Men05_03 
         Caption         =   "&Adelantos de Remuneraciones"
         Begin VB.Menu Men05_03_01 
            Caption         =   "&Adelantos Por Mes"
         End
         Begin VB.Menu Men05_03_02 
            Caption         =   "&Adelantos pendientes de descontar"
         End
         Begin VB.Menu Men05_03_03 
            Caption         =   "&Adelantos por fecha de ingreso al sistema"
         End
      End
      Begin VB.Menu Men05_04 
         Caption         =   "&Cuentas Corrientes"
         Begin VB.Menu Men05_04_01 
            Caption         =   "&Debitos hechos por mes y concepto"
         End
         Begin VB.Menu Men05_04_02 
            Caption         =   "Debitos hechos por concepto"
            Visible         =   0   'False
         End
         Begin VB.Menu Men05_04_03 
            Caption         =   "&Resumen de Debitos por meses"
         End
         Begin VB.Menu Men05_04_04 
            Caption         =   "&Pendientes General"
         End
         Begin VB.Menu Men05_04_05 
            Caption         =   "&Historico por Trabajador"
         End
         Begin VB.Menu mnupendtrab 
            Caption         =   "&Pendientes por Trabajador"
         End
      End
      Begin VB.Menu Men05_05 
         Caption         =   "&Planilla de Remuneraciones"
      End
      Begin VB.Menu Men05_06 
         Caption         =   "&Quinta Categoria"
         Begin VB.Menu Men05_06_01 
            Caption         =   "&Listado de Retenciones"
         End
         Begin VB.Menu Men05_06_02 
            Caption         =   "&Constancia de Retenciones"
         End
         Begin VB.Menu Men05_06_03 
            Caption         =   "&Resumen de Retenciones"
         End
      End
      Begin VB.Menu Men05_07 
         Caption         =   "&Consolidados"
         Begin VB.Menu Men05_07_01 
            Caption         =   "Personal por Area o Centro de Costo"
         End
         Begin VB.Menu Men05_07_02 
            Caption         =   "Distribucion del Personal por Area y por Modalidad de Contratacion"
         End
         Begin VB.Menu Men05_07_03 
            Caption         =   "Costo Mensual del Personal por Area"
         End
      End
      Begin VB.Menu Men05_08 
         Caption         =   "Recursos Humanos"
         Begin VB.Menu Men05_08_01 
            Caption         =   "Reporte Estudios"
            Visible         =   0   'False
         End
         Begin VB.Menu Men05_08_02 
            Caption         =   "Reporte Idiomas"
            Visible         =   0   'False
         End
         Begin VB.Menu Men05_08_03 
            Caption         =   "Reporte Eventos"
            Visible         =   0   'False
         End
         Begin VB.Menu Men05_07_04 
            Caption         =   "Informacion general de Trabajador"
            Visible         =   0   'False
         End
         Begin VB.Menu vaciorrhh 
            Caption         =   "-"
         End
         Begin VB.Menu Men05_08_05 
            Caption         =   "&Citaciones de Trabajadores"
            Visible         =   0   'False
         End
         Begin VB.Menu Men05_08_07 
            Caption         =   "&Relación de Cumpleaños por meses"
         End
         Begin VB.Menu Men05_08_08 
            Caption         =   "Reportes de &Derechohabientes"
         End
         Begin VB.Menu mnurelacontra 
            Caption         =   "Relacion de Contratos"
         End
      End
      Begin VB.Menu Men05_09 
         Caption         =   "Resumen de Planillas por Trabajador"
      End
      Begin VB.Menu Men05_10 
         Caption         =   "Ficha General del Trabajador"
      End
      Begin VB.Menu mnuresuhorext 
         Caption         =   "&Resumen Ingreso Horas Extras"
      End
   End
   Begin VB.Menu Men06 
      Caption         =   "&Opciones"
      Begin VB.Menu Men06_01 
         Caption         =   "&Configuración"
      End
      Begin VB.Menu Men06_02 
         Caption         =   "&Administración de Accesos"
      End
      Begin VB.Menu Men06_07 
         Caption         =   "Crear A&dministradores"
      End
      Begin VB.Menu Men06_03 
         Caption         =   "&Datos Generales de la Empresa"
      End
      Begin VB.Menu Men06_04 
         Caption         =   "&Seleccionar empresa"
         Shortcut        =   ^E
      End
      Begin VB.Menu Men06_05 
         Caption         =   "-"
      End
      Begin VB.Menu Men06_06 
         Caption         =   "&Procesos Auxiliares"
         Begin VB.Menu Men06_06_01 
            Caption         =   "&Actualización de Información"
         End
         Begin VB.Menu Men06_06_02 
            Caption         =   "&Traslado de Información entre empresas"
         End
         Begin VB.Menu Men06_06_03 
            Caption         =   "&Generar Asientos Contables"
         End
      End
   End
   Begin VB.Menu Men07 
      Caption         =   "&Ventana"
      WindowList      =   -1  'True
   End
   Begin VB.Menu Men08 
      Caption         =   "Ay&uda"
      Begin VB.Menu Men08_01 
         Caption         =   "Acerca de..."
      End
      Begin VB.Menu Men08_02 
         Caption         =   "Gratificaciones"
      End
   End
End
Attribute VB_Name = "MDIPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub men01_11_Click()
    ConfiAdel.Show 1
End Sub
Private Sub men01_12_01_Click()
    frCAR.Show
End Sub
Private Sub men01_12_02_Click()
    frVariables.Show
End Sub
Private Sub men01_12_03_Click()
    frBancos.Show
End Sub
Private Sub men01_12_04_Click()
    frDocum.Show
End Sub

Private Sub men01_12_05_Click()
    frColPL.Show
End Sub
Private Sub men01_12_06_Click()
    frTipTra.Show
End Sub
Private Sub men01_12_07_Click()
    frDataTrab.Show 1
End Sub
Private Sub men01_12_08_Click()
    frSubAreas.Show 1
End Sub
Private Sub men01_12_09_Click()
    frFormulasVac.Show 1
End Sub
Private Sub men01_12_10_Click()
    frFormulasGrati.Show 1
End Sub
Private Sub men01_12_11_Click()
    frFormulasCTS.Show 1
End Sub
Private Sub men01_12_12_Click()
    frFormulasUTIL.Show 1
End Sub
Private Sub men01_12_13_Click()
    MantSCat.Show
End Sub
Private Sub men01_12_14_Click()
    FrTipoEst.Show
End Sub
Private Sub men01_12_16_Click()
    FrEstudios.Show
End Sub
Private Sub Men03_02_02_Click()
    frAddAsistCCosto.Show 1
End Sub
Private Sub BarraEstado_PanelClick(ByVal Panel As MSComctlLib.Panel)
    Select Case Panel.KEY
        Case "Dolar"
            frValor.Show 1
            If Val(VPTAREA) <= 0 Then Exit Sub
            If VPTAREA <> "" Then
                Panel.Text = Format(VPTAREA, "0.000 ")
                SaveSetting App.CompanyName, "Planillas", "Dolar", VPTAREA
                Beep
            End If
        Case "PERIODO"
            Dim RSMESES As New ADODB.Recordset
            RSMESES.Open "SELECT MESACTIVO, NOMBRE FROM MESESACT ORDER BY MESACTIVO", DBSYSTEM, adOpenStatic
            If RSMESES.RecordCount = 0 Then
                MsgBox "No se han encontrado meses en actividad", vbCritical
                Set RSMESES = Nothing
                Exit Sub
            End If
            frmComun.CONECTAR RSMESES
            frmComun.Show 1
            If VGUTIL(1) <> "" Then
                BarraEstado.Panels("PERIODO").Text = Month(RSMESES!MESACTIVO) & "/" & Year(RSMESES!MESACTIVO)
            Else
                Set RSMESES = Nothing
                Exit Sub
            End If
            Set RSMESES = Nothing
    End Select
End Sub
Private Sub Men01_03_Click()
    frAreas.Show
End Sub
Private Sub Men03_05_Click()
    'FrmAdel.Show 1
    FrmAdeldet.Show
End Sub
Private Sub Men01_08_Click()
    frPrgPagos.Show 1
End Sub
Private Sub MDIForm_Load()
    BarraEstado.Panels("Dolar").Text = GetSetting(App.CompanyName, "Planillas", "Dolar", "3.50")
End Sub
Private Sub MDIForm_QueryUnload(CANCEL As Integer, UNLOADMODE As Integer)
    If UNLOADMODE = 0 Then
        If MsgBox("Seguro de Salir del Sistema", vbYesNo + vbQuestion) = vbNo Then
            CANCEL = 1
        Else
            End
        End If
    End If
End Sub
Private Sub Men03_01_Click()
    frMesActv.Show
End Sub

Private Sub Men03_06_Click()
    CalcPlan.Show 1
End Sub

Private Sub Men03_07_02_Click()
    frPrmQC.Show 1
End Sub

Private Sub Men03_07_03_Click()
    'frTrb5ta.Show 1
    FrmMant5ta.Show
End Sub

Private Sub Men03_12_Click()
Dim RSRUBROS As New ADODB.Recordset
    RSRUBROS.Open "CONCEPTOS", DBSYSTEM, adOpenKeyset, adLockOptimistic
    If RSRUBROS.EOF Then Exit Sub
    RSRUBROS.MoveFirst
    Dim X As Long, Z As Byte
    X = 0
    Do While Not RSRUBROS.EOF
        If Trim(RSRUBROS!COLPLANILLA) <> "" Then
            DBSYSTEM.Execute "UPDATE COLUMPL SET TIPO=TIPO WHERE CODIGO='" & Trim(RSRUBROS!COLPLANILLA) & "'", X
            If X = 0 Then
                Z = MsgBox("EL CONCEPTO DE REMUNERACIÓN " & RSRUBROS!NOMBRE & " PRESENTA COMO COLUMNA DE PLANILLA EL CÓDIGO " & RSRUBROS!COLPLANILLA & " EL CUAL NO EXISTE DENTRO DE LA BASE DE DATOS. DESEA DEPURAR EL CONCEPTO DE REMUNERACIÓN", vbQuestion + vbYesNoCancel)
                If Z = vbCancel Then Exit Sub
                If Z = vbYes Then
                    VPTAREA = "EDITAR"
                    VPCODTMP = RSRUBROS!Codigo
                    frECnpt.Show 1
                End If
            End If
        End If
        RSRUBROS.MoveNext
    Loop
    MsgBox "LA PLANILLA Y LOS ULTIMOS MOVIMIENTOS SE HAN COMPROBADO Y DAN COMO RESULTADO UN TRABAJO SATISFACTORIO", vbInformation
    Set RSRUBROS = Nothing
End Sub

Private Sub Men03_13_Click()
    frAdminUtil.Show
End Sub

Private Sub Men03_14_Click()
    'FrmProvisiones.Show
    frAdminProvision.Show
End Sub

Private Sub Men03_15_Click()
    CierreMes.Show 1
End Sub

Private Sub Men03_16_Click()
    ReaperturaMes.Show 1
End Sub

Private Sub Men05_05_Click()
    Screen.MousePointer = 11
    frBolEmit.CMFORMATOPLANILLA_Click
    Screen.MousePointer = 1
End Sub
Private Sub Men05_06_01_Click()
CambiaPanelBD True
Screen.MousePointer = 11
    With RpPlanCab
        .Reset
        .ReportFileName = REGSISTEMA.REPORTES & "Plan0077.rpt"
        .Connect = "DSN=" & VGL_SERVERREP & ";User=sa;PWD=;DSQ=planilla_pp"
        .StoredProcParam(0) = REGSISTEMA.BASESQL & ".dbo.TRABAJADORES"
        .StoredProcParam(1) = REGSISTEMA.BASESQL & ".dbo.HIST5TA"
        .StoredProcParam(2) = "CODTRAB"
        .StoredProcParam(3) = "CODTRAB"
        .Formulas(0) = "xEmpresa='" & REGSISTEMA.EMPRESA & "'"
        .Formulas(1) = "xAnno='Año " & Year(Date) & "'"
        .WindowTitle = "Plan0077 - Listado de Retenciones de Quinta Categoria"
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .WindowShowPrintBtn = True
        .WindowShowRefreshBtn = True
        .WindowShowSearchBtn = True
        .WindowShowPrintSetupBtn = True
        If .Status <> 2 Then .Action = 1
        .Reset
    End With
CambiaPanelBD False
Screen.MousePointer = 1
End Sub
Private Sub Men05_06_02_Click()
    With RpPlanCab
        .Reset
        .Connect = "DSN=" & VGL_SERVERREP & ";User=sa;PWD=;DSQ=" & VGL_BASE & ""
        .ReportFileName = REGSISTEMA.REPORTES & "\Plan0076.rpt"
        .Formulas(0) = "xEmpresa='" & REGSISTEMA.EMPRESA & "'"
        .Formulas(1) = "xRuc='" & REGSISTEMA.RUC & "'"
        .Formulas(2) = "xDireccion='" & DevuelveValor("SELECT Dirección FROM Empresa", DBSYSTEM) & " - " & DevuelveValor("SELECT Distrito FROM Empresa", DBSYSTEM) & ", " & DevuelveValor("SELECT Provincia FROM Empresa", DBSYSTEM) & "'"
        .WindowTitle = "Plan0076 - constancias de Retenciones de Quinta Categoria"
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .WindowShowPrintBtn = True
        .WindowShowRefreshBtn = True
        .WindowShowSearchBtn = True
        .WindowShowPrintSetupBtn = True
        If .Status <> 2 Then .Action = 1
        .Reset
    End With
End Sub
Private Sub Men05_06_03_Click()
    With RpPlanCab
        .Reset
        .Connect = "DSN=" & VGL_SERVERREP & ";User=sa;PWD=;DSQ=" & VGL_BASE & ""
        .ReportFileName = REGSISTEMA.REPORTES & "Plan0078.rpt"
        .Formulas(0) = "xEmpresa='" & REGSISTEMA.EMPRESA & "'"
        .WindowTitle = "Plan0078 - Resumen de Retenciones de Quinta Categoria"
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .WindowShowPrintBtn = True
        .WindowShowRefreshBtn = True
        .WindowShowSearchBtn = True
        .WindowShowPrintSetupBtn = True
        If .Status <> 2 Then .Action = 1
    End With
End Sub
Private Sub Men05_07_04_Click()
    If REGSISTEMA.VALRRHH Then Exit Sub
        Dim Trabajador As String
        Dim RSTRAB As New ADODB.Recordset
        RSTRAB.Open "SELECT CODTRAB, NOMBRES FROM VWTRABAJ", DBSYSTEM, adOpenKeyset, adLockReadOnly
        If RSTRAB.EOF Or RSTRAB.RecordCount = 0 Then
            MsgBox "No se han encontrado registro de trabajadores", vbCritical
            Set RSTRAB = Nothing
            Exit Sub
        End If
        frmComun.CONECTAR RSTRAB
        frmComun.Show 1
        If VGUTIL(1) <> "" Then
            Trabajador = RSTRAB!CODTRAB
            With CR1
                .Reset
                .Connect = "DSN=" & VGL_SERVER & ";User=sa;PWD=;DSQ=" & REGSISTEMA.BASESQL & ""
                .WindowTitle = "PlRH0013.rpt - " & Me.Caption
                .ReportFileName = REGSISTEMA.REPORTES & "\PlRH0013.rpt"
                .SelectionFormula = "{Trabajador.Codtrab}='" & Trabajador & "'"
                .Destination = crptToWindow
                .WindowState = crptMaximized
                .WindowShowPrintBtn = True
                .WindowShowRefreshBtn = True
                .WindowShowSearchBtn = True
                .WindowShowPrintSetupBtn = True
                .Formulas(0) = "xEmpresa='" & REGSISTEMA.EMPRESA & "'"
                If .Status <> 2 Then .Action = 1
            End With
        End If
        Set RSTRAB = Nothing
End Sub
Private Sub Men05_07_05_Click()
    If REGSISTEMA.VALRRHH Then frmRRHH001.Show Else MsgBox "No presenta licencia para uso de esta característica del sistema", vbInformation
End Sub
Private Sub Men05_07_07_Click()
    frRelacCumple.Show
End Sub
Private Sub Men05_07_08_Click()
    If REGSISTEMA.VALRRHH Then frmRRHH002.Show Else MsgBox "No presenta licencia para uso de esta característica del sistema", vbInformation
End Sub

Private Sub Men05_08_01_Click()
    Fr_Rep_Est.Show
End Sub

Private Sub Men05_08_02_Click()
     Frm_EXI.Show
End Sub

Private Sub Men05_08_03_Click()
    Fr_Rep_Eve.Show
End Sub

Private Sub Men05_08_07_Click()
    frRelacCumple.Show
End Sub

Private Sub Men05_09_Click()
    'Form2.Show
    frResumenPLTrab.Show
End Sub

Private Sub Men06_06_02_Click()
    VGLFRM = 1
    frEmpTr.Show 1
End Sub

Private Sub Men06_06_03_Click()
    frmGenAsientos.Show 1
End Sub

Private Sub Men06_07_Click()
    frAccesos.Show
End Sub

Private Sub Men08_01_Click()
    frAcerca.Show 1
End Sub
Private Sub Men01_07_Click()
    frCuentas.Show
End Sub
Private Sub Men03_02_07_Click()
    'AQUI PROGRAMAR LA IMPORTACION DEL RELOJ AUTOMATICO ALTERNATIVO...
    'UNA EMPRESA PUEDE TENER MAS DE UN RELOJ AUTOMATICO Y DE DIFERENTE TIPO
End Sub
Private Sub Men05_02_03_Click() 'Consolidado de asistencia
    Load frRngFch
    frRngFch.xFechaFin.Value = Date
    frRngFch.xFechaIni.Value = Date
    frRngFch.xFechaIni.Day = 1
    frRngFch.Label4.Visible = False
    frRngFch.xCampo.Visible = False
    VPTAREA = "PLAN0005.RPT"
    frRngFch.Show 1
End Sub

Private Sub Men06_01_Click()
    frSetup.Show 1
End Sub

Private Sub Men03_02_06_Click()
    MsgBox "No se ha encontrado la libreria del Marcador de Asistencia Automático, Modelo AC-Clock 98' Marca HourExpert. Verifique la instalación del Software de Instalación", vbInformation
End Sub

Private Sub Men05_04_04_Click()
    FrmdebPend.Show 1
End Sub

Private Sub Men06_03_Click()
    frDatos.Show 1
End Sub

Private Sub Men05_02_02_Click()
    Load frRngFch
    frRngFch.xFechaFin.Value = Date
    frRngFch.xFechaIni.Value = Date
    frRngFch.xFechaIni.Day = 1
    VPTAREA = "PLAN0003.RPT"
    frRngFch.Show 1
End Sub

Private Sub Men05_02_01_Click()
    Load frRngFch
    frRngFch.xFechaFin.Value = Date
    frRngFch.xFechaIni.Value = Date
    frRngFch.xFechaIni.Day = 1
    frRngFch.Label3.Caption = "Trabajador"
    frRngFch.Label4.Visible = False
    frRngFch.xCampo.Visible = False
    VPTAREA = "PLAN0004.RPT"
    frRngFch.Show 1
End Sub

Private Sub Men03_03_Click()
    frIngMov.Show 1
End Sub

Private Sub Men06_02_Click()
'    frAccesos.Show 1
frmCrearUsuarios.Show 1
End Sub


Private Sub Men01_05_Click()
    frAFPs.Show
End Sub


Private Sub Men01_04_Click()
     frCostos.Show
End Sub

Private Sub Men01_06_Click()
    frConcpt.Show
End Sub
Private Sub Men01_02_Click()
    frFamily.Show
End Sub


Private Sub Men01_13_Click()
    If MsgBox("Seguro de salir del sistema", vbYesNo + vbQuestion) = vbNo Then Exit Sub
    End
End Sub

Private Sub men01_01_Click()
    frPersonal.Show
End Sub

Private Sub Men02_06_Click()
    On Error GoTo ErrToolBar
    Screen.ActiveForm.COMANDOTOOLBAR "Nuevo"
ErrToolBar:
    Exit Sub
End Sub

Private Sub Men02_08_Click()
    On Error GoTo ErrToolBar
    Screen.ActiveForm.COMANDOTOOLBAR "Buscar"
ErrToolBar:
    Exit Sub
End Sub
Private Sub Men02_07_Click()
    On Error GoTo ErrToolBar
    Screen.ActiveForm.COMANDOTOOLBAR "Eliminar"
ErrToolBar:
    Exit Sub
End Sub
Private Sub Men02_05_Click()
    On Error GoTo ErrToolBar
    Screen.ActiveForm.COMANDOTOOLBAR "Editar"
ErrToolBar:
    Exit Sub
End Sub
Private Sub Men04_02_Click()
    MsgBox "No se ha fijado un Servidor de Almacen de datos en deshuso", vbCritical
End Sub
Private Sub Men04_01_Click()
    frPlans.Show
End Sub

Private Sub Men03_04_Click()
    frAdelantos.Show 1
 '   FrmAdeldet.Show
End Sub
Private Sub Men04_07_Click()
    frFormatos.Show 1
End Sub
Private Sub Men04_04_Click()
    frBolEmit.Show
End Sub
Private Sub Men05_01_Click()
    frAdelPen.Show
End Sub

Private Sub Men04_08_Click()
    Dim RSAUX As New ADODB.Recordset
    Dim STRFILE As String
    STRFILE = DevNomRep(Trim(REGSISTEMA.RUC), Trim(REGSISTEMA.USER), FILEPLANCAB)
    If STRFILE = "" Then
        MsgBox "No se encuentra el NOMBRE del reporte de cabaceras " & Chr(13) & _
               "por favor revise opcion de  configuracion en el sistema", vbExclamation
        Exit Sub
    End If
    
    If UCase(Dir$(REGSISTEMA.REPORTES & STRFILE)) <> UCase(STRFILE) Then
        MsgBox "No se encuentra el NOMBRE del reporte a imprimir"
        Exit Sub
    End If
    Screen.MousePointer = 11
    With RpPlanCab
        .ReportFileName = REGSISTEMA.REPORTES & STRFILE
        .Connect = "DSN=" & VGL_SERVER & ";User=sa;PWD=;DSQ=" & REGSISTEMA.BASESQL & ""
        .Destination = crptToWindow
        .WindowShowPrintBtn = True
        .WindowShowRefreshBtn = True
        .WindowShowSearchBtn = True
        .WindowShowPrintSetupBtn = True
        .WindowState = crptMaximized
        .WindowTitle = STRFILE
        .Formulas(0) = "xEmpresa='" & REGSISTEMA.EMPRESA & "'"
        .Formulas(1) = "xRuc='RUC N° " & REGSISTEMA.RUC & "'"
        '.Formulas(2) = "xMes='Correspondiente al mes de " & LPlans.SelectedItem.Text & "'"
        .Formulas(3) = "xDireccion='" & DevuelveValor("SELECT Dirección FROM Empresa", DBSYSTEM) & "'"
        If .Status <> 2 Then .Action = 1
   End With
   Screen.MousePointer = 1
End Sub

Private Sub Men01_10_Click()
    frmAsientos.Show 1
End Sub

Private Sub Men06_06_01_Click()
    ActualizarSistema
    DBSYSTEM.Execute "UPDATE TRABAJADORES SET FECHACESE=NULL WHERE SITUACIÓN IN ('0','1')"
    DBSYSTEM.Execute "UPDATE TRABAJADORES SET RUCEPS='' WHERE RUCEPS IS NULL"
    DBSYSTEM.Execute "UPDATE CONCEPTOS SET CRITERIO='' WHERE CRITERIO IS NULL"
End Sub

Private Sub Men05_04_05_Click()
    'FrmDebHist.Show 1
    FrmCtahist.Show 1
End Sub

Private Sub Men05_04_01_Click()
    FrmDebMes.Show 1
End Sub

Private Sub Men05_04_03_Click()
    FrmResuDeb.Show 1
End Sub

Private Sub Men03_02_05_Click()
 '   MsgBox "No se ha encontrado la libreria del Marcador de Asistencia Automático, Modelo AC-Clock 98' Marca HourExpert. Verifique la instalación del Software de Instalación", vbInformation
 FrmAsisElect.Show 1
End Sub


Private Sub Men08_02_Click()
    frHelpTemas.Show
End Sub

Private Sub Men05_03_01_Click()
    VPTAREA = "Mes"
    RptAdel.Show 1
End Sub

Private Sub Men04_05_Click()
    frAdelEmit.Show
End Sub

Private Sub Men05_03_03_Click()
    VPTAREA = "Ingreso"
    RptAdel.Show 1
End Sub

Private Sub Men05_03_02_Click()
    VPTAREA = "Pendientes"
    RptAdel.Show 1
End Sub

Private Sub Men03_02_03_Click()
    frTrabajCCostos.Show
End Sub

Private Sub Men03_02_01_Click()
    frRegAsi.Show 1
End Sub

Private Sub Men03_10_Click()
    frAdminLiquid.Show
End Sub

Private Sub Men03_08_Click()
    frVacaciones.Show
End Sub

Private Sub Men03_09_Click()
    frAdminGrati.Show
End Sub

Private Sub Men03_11_Click()
    frAdminCTS.Show
End Sub

Private Sub Men06_04_Click()
    On Error Resume Next
    If Screen.ActiveForm.Name = "MDIPrincipal" Then
        frPanEmp.Show 1
    Else
        MsgBox "Deberá primero cerrar todos los formularios abiertos", vbCritical
        Exit Sub
    End If
    Set DBSYSTEM = New ADODB.Connection
    Set DBADMINPER = Nothing
   With DBSYSTEM 'Para seguridad
       .CursorLocation = adUseClient
       .ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=SOPORTE;Password=SOPORTE;Initial Catalog=" & UCase(REGSISTEMA.BASESQL) & ";Data Source=" & UCase(VGL_SERVER)
       .Open
    End With
    
    With DBADMINPER
        .CommandTimeout = 50
        .CursorLocation = adUseClient
        .Provider = "Microsoft.Jet.OLEDB.3.51"
        .ConnectionString = "Data Source=" & REGSISTEMA.PATHEMPRESA & "\AdminPer.mdb"
        If UCase(Dir$(REGSISTEMA.PATHEMPRESA & "\ADMINPER.MDB")) <> "ADMINPER.MDB" Then REGSISTEMA.VALRRHH = False Else REGSISTEMA.VALRRHH = True
    End With
    
    If REGSISTEMA.VALRRHH Then DBADMINPER.Open
   
    ActualizarSistema
    MDIPrincipal.Caption = "Planillas: " & REGSISTEMA.EMPRESA
End Sub

Private Sub menu01_01_03_Click()
frmBilletes.Show
End Sub

Private Sub mnupendtrab_Click()
    FrmRepCtahistprog.Show
End Sub

Private Sub mnurelacontra_Click()
    frmrelacontra.Show 1
End Sub

Private Sub mnuresuhorext_Click()
    FrmRepHorExt.Show
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error GoTo ErrToolBar
    Select Case UCase(Button.KEY)
        Case "NUEVO", "EDITAR", "ELIMINAR", "IMPRIMIR", "PRELIMINAR", "BUSCAR", "FILTRAR":
            If Men02.Enabled Then Screen.ActiveForm.COMANDOTOOLBAR Button.KEY
        Case "TRABAJADORES": If men01_01.Enabled Then frPersonal.Show
        Case "CALCULADORA": frmCalc.Show
        Case "CCOSTOS": If men01_04.Enabled Then frCostos.Show
        Case "ACTIVAR": If Men03_01.Enabled Then frMesActv.Show
        Case "ADELANTO": If Men03_04.Enabled Then frAdelantos.Show 1
        Case "SALIR": If MsgBox("REALMENTE DESEA SALIR DEL SISTEMA", vbQuestion + vbYesNo) = vbYes Then End
        Case "FORMATOS": If Men04_07.Enabled Then frFormatos.Show 1
        Case "CONCEPTOS": If men01_06.Enabled Then frConcpt.Show
        Case "INPUTBOX": If Men03_05.Enabled Then CalcPlan.Show 1
        Case "PLANILLAS": If Men04_01.Enabled Then frPlans.Show
        Case "SETUP": If Men06_01.Enabled Then frSetup.Show 1
        Case Else:
    End Select
    
    Exit Sub
ErrToolBar:
    Resume Next
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Select Case ButtonMenu.KEY
        Case "IMPOPROD"
                            'MsgBox "Importar"
                    On Error GoTo handler
                    Call ESCRIBIR_WENPLAEXP
                    Shell App.PATH & "\ImportacionExel.exe", vbNormalFocus
End Select

Exit Sub
handler:
    MsgBox "El Sistema de Importación no esta instalado.", vbOKOnly
End Sub
