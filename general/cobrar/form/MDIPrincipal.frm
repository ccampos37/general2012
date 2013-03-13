VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.MDIForm MDIPrincipal 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Sistema de Cobrar"
   ClientHeight    =   6270
   ClientLeft      =   165
   ClientTop       =   525
   ClientWidth     =   10755
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   9000
      Top             =   4920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10755
      _ExtentX        =   18971
      _ExtentY        =   1111
      ButtonWidth     =   609
      ButtonHeight    =   953
      Appearance      =   1
      _Version        =   393216
   End
   Begin Crystal.CrystalReport CrystalReport11 
      Left            =   4680
      Top             =   2760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport CryRptProc 
      Left            =   4920
      Top             =   3480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport oCrystalReport 
      Left            =   4170
      Top             =   7290
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
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
            Picture         =   "MDIPrincipal.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":0452
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":05AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":0716
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":0872
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":09CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":0B2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":0C86
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":0DE6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu opc1 
      Caption         =   "Ingresos"
      Begin VB.Menu opc1_01 
         Caption         =   "Tipo de Planillas"
         Begin VB.Menu opc1_01_01 
            Caption         =   "Planilla Aplicaciones"
            Begin VB.Menu opc1_01_01_01 
               Caption         =   "Documentos"
            End
            Begin VB.Menu opc1_01_01_02 
               Caption         =   "Eliminar"
            End
         End
         Begin VB.Menu opc1_01_02 
            Caption         =   "Varios"
            Begin VB.Menu opc1_01_02_01 
               Caption         =   "Documentos"
            End
            Begin VB.Menu opc1_01_02_02 
               Caption         =   "Eliminar"
            End
         End
         Begin VB.Menu opc1_01_03 
            Caption         =   "Notas Contables Documentario"
            Visible         =   0   'False
            Begin VB.Menu opc1_01_03_01 
               Caption         =   "Documentos"
            End
            Begin VB.Menu opc1_01_03_02 
               Caption         =   "Anulacion"
            End
            Begin VB.Menu opc1_01_03_03 
               Caption         =   "Eliminar"
            End
         End
         Begin VB.Menu opc1_01_04 
            Caption         =   "Canje Renovacion"
            Begin VB.Menu opc1_01_04_01 
               Caption         =   "Canjes"
            End
            Begin VB.Menu opc1_01_04_02 
               Caption         =   "Renovacion"
            End
         End
         Begin VB.Menu opc1_01_05 
            Caption         =   "Nota Abono/Cargo Fisico"
            Visible         =   0   'False
         End
         Begin VB.Menu opc1_01_06 
            Caption         =   "Planilla de Ajustes"
            Begin VB.Menu opc1_01_06_01 
               Caption         =   "Positivo"
            End
            Begin VB.Menu opc1_01_06_02 
               Caption         =   "Negativo"
            End
         End
      End
      Begin VB.Menu opc1_02 
         Caption         =   "Tablas"
         Begin VB.Menu opc1_02_01 
            Caption         =   "Bancos"
            Visible         =   0   'False
         End
         Begin VB.Menu opc1_02_02 
            Caption         =   "Tipo Documentos"
         End
         Begin VB.Menu opc1_02_03 
            Caption         =   "Concepto"
         End
         Begin VB.Menu opc1_02_04 
            Caption         =   "Vendedores"
         End
         Begin VB.Menu opc1_02_05 
            Caption         =   "Empresa"
         End
         Begin VB.Menu opc1_02_06 
            Caption         =   "Zonas de Ventas"
            Visible         =   0   'False
         End
         Begin VB.Menu opc1_02_07 
            Caption         =   "Tipo de Negocio"
            Visible         =   0   'False
         End
         Begin VB.Menu opc1_02_08 
            Caption         =   "Tipo de Planillas"
         End
         Begin VB.Menu opc1_02_09 
            Caption         =   "Limite de Credito"
            Visible         =   0   'False
            Begin VB.Menu opc1_02_09_01 
               Caption         =   "Grupo Empresarial"
            End
            Begin VB.Menu opc1_02_09_02 
               Caption         =   "x Tipo Documento"
            End
         End
      End
      Begin VB.Menu opc1_03 
         Caption         =   "Maestros"
         Begin VB.Menu opc1_03_01 
            Caption         =   "Clientes"
         End
      End
   End
   Begin VB.Menu opc2 
      Caption         =   "Procesos"
      Begin VB.Menu opc2_01 
         Caption         =   "Cierre Mensuak"
         Visible         =   0   'False
      End
      Begin VB.Menu opc2_02 
         Caption         =   "Modificacion Documentos"
         Visible         =   0   'False
      End
      Begin VB.Menu opc2_03 
         Caption         =   "Regeneracion de Saldos"
      End
      Begin VB.Menu opc2_04 
         Caption         =   "Anulacion Letras"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu opc3 
      Caption         =   "Reportes"
      Begin VB.Menu opc3_01 
         Caption         =   "Saldos"
         Begin VB.Menu opc3_01_01 
            Caption         =   "Clientes"
         End
         Begin VB.Menu opc3_01_02 
            Caption         =   "Vendedores"
            Visible         =   0   'False
         End
      End
      Begin VB.Menu opc3_02 
         Caption         =   "Estado Cta Cte"
         Begin VB.Menu opc3_02_01 
            Caption         =   "Clientes"
         End
         Begin VB.Menu opc3_02_02 
            Caption         =   "Vendedores"
         End
      End
      Begin VB.Menu opc3_03 
         Caption         =   "Planilla Cobranza"
      End
      Begin VB.Menu opc3_04 
         Caption         =   "Planilla Varios"
      End
      Begin VB.Menu opc3_05 
         Caption         =   "Resumen  Cobranza"
         Visible         =   0   'False
         Begin VB.Menu opc3_05_01 
            Caption         =   "Diario"
         End
         Begin VB.Menu opc3_05_02 
            Caption         =   "detallado"
         End
      End
      Begin VB.Menu opc3_06 
         Caption         =   "Documentos"
         Visible         =   0   'False
         Begin VB.Menu opc2_06_01 
            Caption         =   "Listado General"
         End
         Begin VB.Menu opc3_06_02 
            Caption         =   "Vencidos/x vencer"
         End
         Begin VB.Menu opc3_06_03 
            Caption         =   "x Compensar"
         End
         Begin VB.Menu opc3_06_04 
            Caption         =   "Vencido/x vencer(formato 2)"
         End
         Begin VB.Menu opc3_06_05 
            Caption         =   "Aviso de Cobranza"
         End
      End
      Begin VB.Menu opc3_07 
         Caption         =   "Notas Contables"
         Visible         =   0   'False
      End
      Begin VB.Menu opc3_08 
         Caption         =   "Clientes"
      End
      Begin VB.Menu opc3_09 
         Caption         =   "Planillas de Canjes"
         Visible         =   0   'False
         Begin VB.Menu opc3_09_01 
            Caption         =   "Tipo de Planillas"
         End
         Begin VB.Menu opc3_09_02 
            Caption         =   "Documentos canjeados"
         End
      End
      Begin VB.Menu opc3_a 
         Caption         =   "Planilla de renovaciones"
         Visible         =   0   'False
      End
      Begin VB.Menu opc3_c 
         Caption         =   "Letras"
         Visible         =   0   'False
         Begin VB.Menu opc3_c_01 
            Caption         =   "Descontadas"
            Visible         =   0   'False
         End
         Begin VB.Menu opc3_c_02 
            Caption         =   "Impresion"
         End
      End
      Begin VB.Menu menu03_11 
         Caption         =   "Planilla Bancos"
         Visible         =   0   'False
      End
      Begin VB.Menu opc3_B 
         Caption         =   "Modo Venta"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu opc4 
      Caption         =   "Consultas"
      Begin VB.Menu opc4_01 
         Caption         =   "Saldos x Clientes"
      End
   End
   Begin VB.Menu opc5 
      Caption         =   "Configuracion"
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
Option Explicit

Private Sub Form_Load()
 Unload FrmIngreso
 MostrarForm Me, "M"
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If MsgBox("Desea Salir del Sistema?", vbYesNo, "AVISO") = vbYes Then
      Set VGGeneral = Nothing
      Set VGCNx = Nothing
      Set VGcnxCT = Nothing
      End
   End If
End Sub

Private Sub menu03_11_Click()
    FrmRepPlanillaCob_Banco.Show
End Sub

Private Sub mnuavicobra_Click()

End Sub

Private Sub mnuDocxgrupcred_Click()
   frmDocxLimit.Show
End Sub

Private Sub mnugruplimicred_Click()
   fmrlimitgrupo.Show
End Sub

Private Sub Opc1_01_01_01_Click()
VGPlanillaAjuste = 0
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

Private Sub Opc1_01_03_04_Click()
   
End Sub

Private Sub Opc1_01_04_01_Click()
  FrmPlanillaCanjes.Show
End Sub

Private Sub Opc1_01_04_02_Click()
  FrmPlanillaRenova.Show
End Sub

Private Sub Opc1_01_05_Click()
  FrmNotaFisico.Show
End Sub

Private Sub opc1_01_06_01_Click()
VGPlanillaAjuste = 1
  FrmPlanillaCobranza.Show
End Sub


Private Sub opc1_01_06_02_Click()
  VGPlanillaAjuste = 2
  FrmPlanillaCobranza.Show
End Sub



Private Sub Opc1_02_01_Click()
  frmBanco.Show
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

Private Sub Opc1_02_07_Click()
 FrmNegocio.Show
End Sub

Private Sub Opc1_02_08_Click()
  FrmTipoPlanilla.Show
End Sub

Private Sub Opc1_03_01_Click()
 Frmcliente.Show
End Sub

Private Sub Opc1_03_02_Click()
  FrmLimiteCredito.Show
End Sub

Private Sub Opc1_03_03_Click()
 FrmMultidireccion.Show
End Sub

Private Sub Opc1_03_04_Click()
   frmClientexGrupoCred.Show
End Sub

Private Sub opc2_03_Click()
  If MsgBox("Desea Regenerar los Saldos?", vbYesNo, MsgTitle) = vbYes Then
     FrmGeneraSaldos.Show 1
  End If
End Sub

Private Sub opc2_04_Click()
   frmAnularLetras.Show
End Sub

Private Sub opc2_05_Click()
 Dim SQL As String
   Screen.MousePointer = 11
   SQL = "insert ct_tipocambio "
   SQL = SQL & "select * from " & g_BaseContab & ".dbo.ct_tipocambio where tipocambiofecha not in"
   SQL = SQL & "(select tipocambiofecha from ct_tipocambio)"
   VGCNx.Execute (SQL)
   Screen.MousePointer = 1
End Sub

Private Sub opc3_01_01_Click()
  FrmSaldoxCliente.Show
End Sub

Private Sub opc3_01_02_Click()
  RptSaldoxVendedor.Show
End Sub

Private Sub opc3_02_01_Click()
 RptctactexCliente.Show
End Sub

Private Sub opc3_02_02_Click()
    RptCtactexVendedor.Show
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

Private Sub opc3_06_01_Click()
  frmRepListadoDocumentos.Show
End Sub

Private Sub opc3_06_02_Click()
  RptDocumentosxCobrar.Show
End Sub

Private Sub opc3_06_03_Click()
  RptDocumentosxAplicar.Show
End Sub

Private Sub opc3_06_04_Click()
 FrmRepDocvenciXvence.Show
End Sub

Private Sub opc3_06_05_Click()
    frmrepavisocobra.Show
End Sub

Private Sub opc3_07_Click()
  RptNotaabono.Show
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

Private Sub opc3_08_Click()
   frmRepClientes.Show
End Sub

Private Sub opc3_09_01_Click()
  FrmRepPlanillaCanjeRenovacion.Opcion = "1"
  FrmRepPlanillaCanjeRenovacion.Show
End Sub


Private Sub opc3_D_Click()
   frmrepantigdeudas.Show
End Sub

Private Sub opc3_03_02_Click()
  FrmRepOtroPlanillaCanjeRenovacion.Opcion = "1"
  FrmRepOtroPlanillaCanjeRenovacion.Show
End Sub

Private Sub opc3_A_Click()
  FrmRepOtroPlanillaCanjeRenovacion.Opcion = "2"
  FrmRepOtroPlanillaCanjeRenovacion.Show
End Sub

Private Sub opc3_B_Click()
  RptModoVenta.Show
End Sub

Private Sub opc3_C_01_Click()
  ' frmRepLetrasDescontadas.Show
End Sub

Private Sub opc3_C_02_Click()
  frmRepImpresionLetras.Show
End Sub

Private Sub opc4_01_Click()
  CstSaldoCliente.Show
End Sub

Private Sub opc5_01_Click()
VGtipo = cobrar
' frmCfgUsuario.Show 1
End Sub

Private Sub opc6_Click()
   If MsgBox("Desea Salir del Sistema?", vbYesNo, "AVISO") = vbYes Then
      Set cn = Nothing
      Set VGGeneral = Nothing
      Set VGCNx = Nothing
      Set VGcnxCT = Nothing
      End
   End If
End Sub



