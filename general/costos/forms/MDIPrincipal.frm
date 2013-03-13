VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.MDIForm MDIPrincipal 
   BackColor       =   &H8000000B&
   Caption         =   "MDIForm1"
   ClientHeight    =   8190
   ClientLeft      =   165
   ClientTop       =   360
   ClientWidth     =   11880
   Icon            =   "MDIPrincipal.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CryRptProc 
      Left            =   5280
      Top             =   3720
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
      Top             =   7860
      Width           =   11880
      _ExtentX        =   20955
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
            Picture         =   "MDIPrincipal.frx":0442
            Text            =   "Tipo Cambio"
            TextSave        =   "Tipo Cambio"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4480
            MinWidth        =   4480
            Picture         =   "MDIPrincipal.frx":075E
            Text            =   "Servidor"
            TextSave        =   "Servidor"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7832
            MinWidth        =   7832
            Picture         =   "MDIPrincipal.frx":08BA
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
   Begin VB.Menu menu01 
      Caption         =   "tablas"
   End
   Begin VB.Menu Menu04 
      Caption         =   "Procesos"
      Begin VB.Menu menu04_01 
         Caption         =   "Resumen mensual"
      End
      Begin VB.Menu menu04_02 
         Caption         =   "control de Cierres"
      End
   End
   Begin VB.Menu Menu03 
      Caption         =   "Reportes"
      Begin VB.Menu menu03_01 
         Caption         =   "Resumenes"
         Begin VB.Menu menu03_01_01 
            Caption         =   "Diario"
         End
         Begin VB.Menu menu03_01_02 
            Caption         =   "Mes"
         End
         Begin VB.Menu Menu03_01_03 
            Caption         =   "Punto de equilibrio"
         End
      End
      Begin VB.Menu menu03_02 
         Caption         =   "Estadisticas"
         Begin VB.Menu menu03_02_01 
            Caption         =   "Mensualizada"
         End
         Begin VB.Menu menu03_02_02 
            Caption         =   "Grafica"
         End
         Begin VB.Menu menu03_02_03 
            Caption         =   "Costo Unitario Mensualizado"
         End
         Begin VB.Menu menu03_02_04 
            Caption         =   "Costo unitario x Dia del mes"
         End
      End
      Begin VB.Menu menu03_04 
         Caption         =   "Costos 2012"
         Begin VB.Menu menu03_04_01 
            Caption         =   "Resumen mensual"
         End
         Begin VB.Menu menu03_04_02 
            Caption         =   "Resumen por Centro de Costos"
         End
      End
   End
   Begin VB.Menu Menu05 
      Caption         =   "Configuracion"
      Begin VB.Menu Menu05_01 
         Caption         =   "Creacion de usuarios"
      End
   End
   Begin VB.Menu menuSalir 
      Caption         =   "Salir"
   End
End
Attribute VB_Name = "MDIPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
If VGCNx.State = 1 Then VGCNx.Close
If VGcnxCT.State = 1 Then VGcnxCT.Close
'Unload Me
End
End Sub

Private Sub menu03_01_01_Click()
FrmresumenDiario.Show 1
End Sub

Private Sub menu03_01_02_Click()
FrmResumenGeneral.Show 1
End Sub

Private Sub menu03_01_03_Click()
FrmPuntoEquilibrio.Show 1
End Sub

Private Sub menu03_02_01_Click()
FrmresumenesMensuales.Show 1
End Sub

Private Sub menu03_02_02_Click()
FrmResumenesMensualesGrafica.Show 1
End Sub

Private Sub menu03_02_03_Click()
FrmCostoUnitarioxMeses.Show 1
End Sub

Private Sub menu03_02_04_Click()
FrmCostoxdiaxmes.Show 1
End Sub

Private Sub menu03_04_01_Click()
FrmResumeGeneralReporte.Show 1
End Sub

Private Sub menu04_01_Click()
FrmResumengeneralNuevo.Show 1
End Sub

Private Sub menu03_04_02_Click()
FrmresumenxcentroCosto.Show 1
End Sub

Private Sub menu04_02_Click()
FrmCierres.Show 1
End Sub

Private Sub Menu05_01_Click()
frmCfgUsuario.Show
End Sub

Private Sub menuSalir_Click()
If VGCNx.State = 1 Then VGCNx.Close
If VGcnxCT.State = 1 Then VGcnxCT.Close
'Unload Me
End
End Sub

Private Sub StatusBar1_PanelClick(ByVal Panel As MSComctlLib.Panel)
    Select Case Panel.Index
        Case 1, 2
            FrmSeleMes.Show 1
    End Select
End Sub
