VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form MDIPrincipal 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Sistema de Cuentas Por Pagar"
   ClientHeight    =   7800
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   11355
   Icon            =   "MDIPrincipal.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   Picture         =   "MDIPrincipal.frx":0CCA
   ScaleHeight     =   7800
   ScaleWidth      =   11355
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar panel 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   7425
      Width           =   11355
      _ExtentX        =   20029
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin Crystal.CrystalReport CryRptProc 
      Left            =   4170
      Top             =   7290
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComctlLib.Toolbar toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11355
      _ExtentX        =   20029
      _ExtentY        =   1164
      ButtonWidth     =   609
      ButtonHeight    =   1005
      Appearance      =   1
      _Version        =   393216
   End
   Begin VB.Menu Opc1 
      Caption         =   "Movimientos"
      Begin VB.Menu Opc1_01 
         Caption         =   "Ingreso Datos"
         Begin VB.Menu Opc1_01_01 
            Caption         =   "Planilla Aplicaciones"
            Begin VB.Menu Opc1_01_01_01 
               Caption         =   "Ingreso Documentos"
            End
            Begin VB.Menu Opc1_01_01_02 
               Caption         =   "Elimina Documentos de Planilla"
            End
         End
         Begin VB.Menu Opc1_01_02 
            Caption         =   "Documentos Varios"
            Begin VB.Menu Opc1_01_02_01 
               Caption         =   "Ingreso Documentos"
            End
            Begin VB.Menu Opc1_01_02_02 
               Caption         =   "Elimina Documentos de Planilla"
            End
         End
         Begin VB.Menu opt1 
            Caption         =   "-"
         End
         Begin VB.Menu Opc1_01_03 
            Caption         =   "Nota Abono/Cargo (No Va en CP)"
            Visible         =   0   'False
            Begin VB.Menu Opc1_01_03_01 
               Caption         =   "Ingresa Documento en Cta. Cte."
            End
            Begin VB.Menu Opc1_01_03_02 
               Caption         =   "Anula Documento Registrado"
            End
            Begin VB.Menu Opc1_01_03_03 
               Caption         =   "Elimina Documento Registrado"
            End
         End
         Begin VB.Menu Opc1_01_05 
            Caption         =   "Nota Abono/Cargo Fisico (No Va en CP)"
            Visible         =   0   'False
         End
         Begin VB.Menu Opc1_01_04 
            Caption         =   "Canje Renovacion"
            Begin VB.Menu Opc1_01_04_01 
               Caption         =   "Canje de Documentos"
            End
            Begin VB.Menu Opc1_01_04_02 
               Caption         =   "Renovacion Documentos"
            End
            Begin VB.Menu Opc1_01_04_03 
               Caption         =   "Anulacion Canjes/renovaciones"
            End
         End
         Begin VB.Menu Opc1_01_06 
            Caption         =   "Planilla Compensaciones"
            Visible         =   0   'False
            Begin VB.Menu Opc1_01_06_01 
               Caption         =   "Ingresa Documentos"
            End
            Begin VB.Menu Opc1_01_06_02 
               Caption         =   "Elimina Documentos"
            End
         End
      End
      Begin VB.Menu Opc1_02 
         Caption         =   "Actualiza Tablas"
         Begin VB.Menu Opc1_02_01 
            Caption         =   "Bancos"
         End
         Begin VB.Menu Opc1_02_02 
            Caption         =   "Tipos Documentos"
         End
         Begin VB.Menu Opc1_02_03 
            Caption         =   "Conceptos"
            Visible         =   0   'False
         End
         Begin VB.Menu Opc1_02_04 
            Caption         =   "Oficinas"
         End
         Begin VB.Menu Opc1_02_05 
            Caption         =   "Empresas"
         End
         Begin VB.Menu Opc1_02_06 
            Caption         =   "Zonas"
            Visible         =   0   'False
         End
         Begin VB.Menu Opc1_02_07 
            Caption         =   "Tipo de Negocio"
         End
         Begin VB.Menu Opc1_02_08 
            Caption         =   "Tipo Planillas"
         End
         Begin VB.Menu Opc1_02_09 
            Caption         =   "Codigo Postal"
         End
         Begin VB.Menu Opc1_02_10 
            Caption         =   "Forma de Pagos"
         End
      End
      Begin VB.Menu Opc1_03 
         Caption         =   "Actualiza Maestros"
         Begin VB.Menu Opc1_03_01 
            Caption         =   "Proveedores"
         End
         Begin VB.Menu Opc132 
            Caption         =   "Limite Credito"
            Visible         =   0   'False
         End
         Begin VB.Menu Opc133 
            Caption         =   "Direcciones Proveedores"
            Visible         =   0   'False
         End
      End
   End
   Begin VB.Menu opc2 
      Caption         =   "Procesos"
      Begin VB.Menu opc2_01 
         Caption         =   "Cierre Mensual"
         Enabled         =   0   'False
      End
      Begin VB.Menu opc2_02 
         Caption         =   "Regularizacion Facturas"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu opc2_03 
         Caption         =   "Regeneracion Saldos"
      End
      Begin VB.Menu opc2_04 
         Caption         =   "Anulacion de Letras"
      End
      Begin VB.Menu opc2_05 
         Caption         =   "Contabilizacion"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu OPC3 
      Caption         =   "Reportes"
      Begin VB.Menu opc3_01 
         Caption         =   "Saldo Documentos"
         Begin VB.Menu opc3_01_01 
            Caption         =   "Saldo por Proveedor"
         End
         Begin VB.Menu opc3_01_02 
            Caption         =   "Saldo por Vendedor"
            Enabled         =   0   'False
            Visible         =   0   'False
         End
      End
      Begin VB.Menu opc3_02 
         Caption         =   "Estado Cta Cte"
         Begin VB.Menu opc3_02_01 
            Caption         =   "Cta Cte x Vendedor"
            Enabled         =   0   'False
            Visible         =   0   'False
         End
         Begin VB.Menu opc3_02_02 
            Caption         =   "Cta Cte x Proveedores"
         End
      End
      Begin VB.Menu opc3_03 
         Caption         =   "Planilla Pagos"
      End
      Begin VB.Menu opc3_04 
         Caption         =   "Planilla Varios"
      End
      Begin VB.Menu opcPlanillaOtros 
         Caption         =   "Planilla Canje/Renovacion"
         Begin VB.Menu opcPlanCanje 
            Caption         =   "Planilla de Canje"
         End
         Begin VB.Menu opcPlanRenovacion 
            Caption         =   "Planilla de Renovación"
         End
      End
      Begin VB.Menu opc3_05 
         Caption         =   "Planilla Compensaciones"
         Visible         =   0   'False
      End
      Begin VB.Menu mnudeuporpag 
         Caption         =   "Resumen Deudas por Proveedor"
         Visible         =   0   'False
      End
      Begin VB.Menu mnudocvencixvencer 
         Caption         =   "Documentos Vencidos x Vencer"
         Visible         =   0   'False
      End
      Begin VB.Menu opc3_06_01 
         Caption         =   "Antiguedad de Deuda"
         Visible         =   0   'False
      End
      Begin VB.Menu OPC3_06_02 
         Caption         =   "documentos pendientes"
         Visible         =   0   'False
      End
      Begin VB.Menu mnureladoc 
         Caption         =   "Relacion de Documentos"
         Visible         =   0   'False
      End
      Begin VB.Menu OPCNOTA 
         Caption         =   "Nota Abono/Cargo"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu opc3_08 
         Caption         =   "Proveedores Reportes"
         Enabled         =   0   'False
         Visible         =   0   'False
         Begin VB.Menu opc3_08_01 
            Caption         =   "Proveedor General"
         End
         Begin VB.Menu opc3_08_02 
            Caption         =   "Proveedor x Zona"
         End
         Begin VB.Menu opc3_08_03 
            Caption         =   "Proveedor x Vendedor"
         End
         Begin VB.Menu opc3_08_04 
            Caption         =   "Proveedor x Distrito"
         End
         Begin VB.Menu opc3_08_05 
            Caption         =   "Proveedor x Categoria"
         End
      End
   End
   Begin VB.Menu opc4 
      Caption         =   "Consultas"
      Begin VB.Menu opc4_01 
         Caption         =   "Saldo por Proveedor"
      End
      Begin VB.Menu opc4_02 
         Caption         =   "Migracion Conta"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu opc5 
      Caption         =   "Configuracion"
      Visible         =   0   'False
      Begin VB.Menu opc5_01 
         Caption         =   "Usuarios"
      End
   End
   Begin VB.Menu opc6 
      Caption         =   "Salir"
   End
End
Attribute VB_Name = "MDIPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    MostrarForm Me, "M"
End Sub
Private Sub Form_Unload(Cancel As Integer)
'  Call opc6_Click
End Sub
Private Sub mnudeuporpag_Click()
    frmrepantigdeudas.Show
End Sub
Private Sub mnudocvencixvencer_Click()
  FrmRepDocvenciXvence.Show
End Sub
Private Sub mnureladoc_Click()
  RptRelaDocumentos.Show
End Sub
Private Sub mnusubtotd_Click()
  RptDocumentosxPagar_2.Show
End Sub
Private Sub Opc1_01_01_01_Click()
  VGaplicaciones = 0
  FrmPlanillaCobranza.Show
End Sub
Private Sub Opc1_01_01_02_Click()
  FrmPlanillaCobranzaModi.Show
End Sub
Private Sub Opc1_01_02_01_Click()
   FrmPlanillaVarios.Show
End Sub
Private Sub Opc1_01_02_02_Click()
  FrmPlanillaVariosModi.Show
End Sub
Private Sub Opc1_01_03_01_Click()
 FrmNotas.Show
End Sub
Private Sub Opc1_01_03_02_Click()
  FrmAnulaNota.Show
End Sub
Private Sub Opc1_01_03_03_Click()
  FrmEliminaNota.Show
End Sub
Private Sub Opc1_01_04_01_Click()
  FrmPlanillaCanjes.Show
End Sub
Private Sub Opc1_01_04_02_Click()
  FrmPlanillaRenova.Show
End Sub

Private Sub Opc1_01_04_03_Click()
 FrmAnulaPllaCanjes.Show 1
End Sub

Private Sub Opc1_01_05_Click()
  FrmNotaFisico.Show
End Sub
Private Sub Opc1_01_06_01_Click()
  VGaplicaciones = 1
  FrmPlanillaCompensaciones.Show
End Sub

Private Sub Opc1_01_06_02_Click()
FrmPlanillaCobranzaModi.Show
End Sub

Private Sub Opc1_02_01_Click()
  frmBanco.Show
End Sub

Private Sub Opc1_02_10_Click()
FrmFormadePago.Show
End Sub

Private Sub Opc1_02_02_Click()
  FrmTipodocumentos.Show
End Sub

Private Sub Opc1_02_03_Click()
  FrmTipoConcepto.Show
End Sub

Private Sub Opc1_02_04_Click()
  FrmVendedor.Show
End Sub

Private Sub Opc1_02_05_Click()
  FrmEmpresa.Show
End Sub

Private Sub Opc1_02_06_Click()
  FrmZona.Show
End Sub

Private Sub Opc127_Click()
 FrmNegocio.Show
End Sub

Private Sub Opc1_02_08_Click()
  FrmTipoPlanilla.Show
End Sub

Private Sub Opc1_02_09_Click()
 frmCodigoPostal.Show
End Sub

Private Sub Opc1_03_01_Click()
 FrmProveedor.Show 1
End Sub
Private Sub Opc1_03_02_Click()
  FrmLimiteCredito.Show
End Sub

Private Sub Opc1_03_03_Click()
 FrmMultidireccion.Show
End Sub

Private Sub opc2_03_Click()
  If MsgBox("Desea Regenerar los Saldos?", vbYesNo, MsgTitle) = vbYes Then
     PrcGeneraSaldos.Show 1
     Unload PrcGeneraSaldos
  End If
End Sub

Private Sub opc2_04_Click()
   frmAnularLetras.Show
End Sub

Private Sub opc2_05_Click()
FrmContabilizacion.Show 1
End Sub

Private Sub opc3_01_01_Click()
   RptSaldoxProveedor.Show
End Sub

Private Sub opc3_01_02_Click()
  RptSaldoxVendedor.Show
End Sub

Private Sub opc3_02_01_Click()
    RptCtactexVendedor.Show
End Sub

Private Sub opc3_02_02_Click()
    RptctactexCliente.Show
End Sub

Private Sub opc3_03_Click()
    FrmRepPlanillaCob.Show
End Sub

Private Sub opc3_04_Click()
    FrmRepPlanillaDocVar.Show
End Sub

Private Sub opc3_05_01_Click()
    RptResumenCobranzaDiaria.Show
End Sub

Private Sub opc3_05_02_Click()
    RptResumenCobranzaDetallada.Show
End Sub

Private Sub opc3_06_Click()
    RptDocumentosxPagar.Show
End Sub

Private Sub opc3_05_Click()
 FrmrepPlanCompensaciones.Show
End Sub

Private Sub opc3_06_02_Click()
  RptDocumentosxPagar.Show
End Sub

Private Sub opc3_06_03_Click()
  RptDocumentosxAplicar.Show
End Sub

Private Sub opc3_07_Click()
  RptDocumentosxPagar_2.Show
End Sub

Private Sub opc3_08_02_Click()
    Rptclientexzona.Show
End Sub

Private Sub opc3_08_03_Click()
   RptclientexVendedor.Show
End Sub

Private Sub opc3_08_04_Click()
   RptClientexdistrito.Show
End Sub

Private Sub opc3_08_05_Click()
  Rptclientexcategoria.Show
End Sub

Private Sub opc4_01_Click()
  CstSaldoCliente.Show
End Sub

Private Sub opc5_01_Click()
VGtipo = pagar
frmCfgUsuario.Show
End Sub

Private Sub opc6_Click()
   If MsgBox("Desea Salir del Sistema?", vbYesNo, "AVISO") = vbYes Then
      Set VGCNx = Nothing
      Set VGgeneral = Nothing
      Set VGCNx = Nothing
      Set VGcnxCT = Nothing
      Unload Me
      End
   End If
   
End Sub

Private Sub opcPlanCanje_Click()
   FrmRepOtroPlanillaCanjeRenovacion.Opcion = "1"
   FrmRepOtroPlanillaCanjeRenovacion.Show
End Sub

Private Sub opcPlanRenovacion_Click()
   FrmRepOtroPlanillaCanjeRenovacion.Opcion = "2"
   FrmRepOtroPlanillaCanjeRenovacion.Show
End Sub


