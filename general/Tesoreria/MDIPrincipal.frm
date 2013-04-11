VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.MDIForm MDIPrincipal 
   BackColor       =   &H8000000C&
   Caption         =   "Sistema de Tesoreria"
   ClientHeight    =   7515
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   12075
   Icon            =   "MDIPrincipal.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CryRptProc 
      Left            =   30
      Top             =   4395
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComctlLib.ImageList ImageList1 
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
            Picture         =   "MDIPrincipal.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":0894
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":09F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":0B58
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":0CB4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":0E10
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":0F6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":10C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":1228
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar ToolComprob 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   18960
      _ExtentX        =   33443
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
   Begin MSComctlLib.StatusBar Panel 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   1
      Top             =   10785
      Width           =   18960
      _ExtentX        =   33443
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3881
            MinWidth        =   3881
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
            Picture         =   "MDIPrincipal.frx":1388
            Text            =   "Tipo Cambio"
            TextSave        =   "Tipo Cambio"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
            Picture         =   "MDIPrincipal.frx":16A4
            Text            =   "Servidor"
            TextSave        =   "Servidor"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11104
            MinWidth        =   4410
            Picture         =   "MDIPrincipal.frx":1800
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
         Caption         =   "Insertar"
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
      End
   End
   Begin VB.Menu opc1 
      Caption         =   "&Movimientos"
      Begin VB.Menu opc11 
         Caption         =   "Registro de Movimientos"
         Begin VB.Menu opc111 
            Caption         =   "Operaciones  Generales"
            Begin VB.Menu opc1111 
               Caption         =   "Varias "
            End
            Begin VB.Menu opc1112 
               Caption         =   "Cuentas por Cobrar"
            End
            Begin VB.Menu opc1113 
               Caption         =   "Cuentas por Pagar"
            End
            Begin VB.Menu opc115 
               Caption         =   "Transferencias"
               Begin VB.Menu opc1151 
                  Caption         =   "Banco a Banco"
               End
               Begin VB.Menu opc1152 
                  Caption         =   "Caja a Banco"
               End
               Begin VB.Menu opc1153 
                  Caption         =   "Banco a Caja"
               End
               Begin VB.Menu opc1154 
                  Caption         =   "Caja a Caja"
               End
            End
         End
         Begin VB.Menu opc114 
            Caption         =   "Conciliaciónes"
            Visible         =   0   'False
            Begin VB.Menu opc1141 
               Caption         =   "Conciliacion Bancos"
            End
            Begin VB.Menu opc1142 
               Caption         =   "Rendiciones de caja"
            End
         End
         Begin VB.Menu opc116 
            Caption         =   "Anulacion de Recibos"
            Begin VB.Menu opc1161 
               Caption         =   "Recibos de Ingreso/Egreso"
            End
            Begin VB.Menu opc1163 
               Caption         =   "Transferencias"
            End
         End
         Begin VB.Menu opc119 
            Caption         =   "Diversos"
            Visible         =   0   'False
            Begin VB.Menu opc11491 
               Caption         =   "Cheques Cobrado en Banco"
               Visible         =   0   'False
            End
            Begin VB.Menu opc1192 
               Caption         =   "Cheques Depositado en Banco"
               Visible         =   0   'False
            End
            Begin VB.Menu opc1193 
               Caption         =   "Giros Cheques Especiales"
               Visible         =   0   'False
            End
            Begin VB.Menu opt1 
               Caption         =   "-"
               Visible         =   0   'False
            End
            Begin VB.Menu opc1195 
               Caption         =   "Registra Cheques Devueltos Proveedor"
               Visible         =   0   'False
            End
            Begin VB.Menu opc1196 
               Caption         =   "Registra Cheques Diferidos Proveedor"
               Visible         =   0   'False
            End
            Begin VB.Menu opt2 
               Caption         =   "-"
            End
            Begin VB.Menu opc1198 
               Caption         =   "Protesta Letras en Descuento Clientes"
               Visible         =   0   'False
            End
            Begin VB.Menu opc1199 
               Caption         =   "Ingreso de Cheques Devueltos Clientes"
               Visible         =   0   'False
            End
         End
         Begin VB.Menu opc117 
            Caption         =   "Documentos x Rendir"
            Begin VB.Menu opc1171 
               Caption         =   "Transferencias"
               Begin VB.Menu opc11711 
                  Caption         =   "Voucher x rendir"
               End
               Begin VB.Menu opc11712 
                  Caption         =   "Recibos x rendir"
               End
            End
            Begin VB.Menu opc1172 
               Caption         =   "Devoluciones "
               Begin VB.Menu opc11721 
                  Caption         =   "Voucher x rendir"
               End
               Begin VB.Menu opc11722 
                  Caption         =   "Recibos x rendir"
               End
            End
            Begin VB.Menu opc1173 
               Caption         =   "Operaciones Varias"
            End
            Begin VB.Menu opc1174 
               Caption         =   "Registrar  Pagos"
            End
            Begin VB.Menu opc1175 
               Caption         =   "Cuentas x pagar"
            End
         End
         Begin VB.Menu opc118 
            Caption         =   "Fondo fijo"
            Visible         =   0   'False
            Begin VB.Menu opc1181 
               Caption         =   "Transferencias"
               Begin VB.Menu opc11811 
                  Caption         =   "Voucher x rendir"
               End
               Begin VB.Menu opc11812 
                  Caption         =   "Recibos x rendir"
               End
               Begin VB.Menu opc11813 
                  Caption         =   "Vales Provisionales"
               End
            End
            Begin VB.Menu opc1182 
               Caption         =   "Devolucion"
               Begin VB.Menu opc11821 
                  Caption         =   "Voucher x rendir"
               End
               Begin VB.Menu opc11822 
                  Caption         =   "Recibos x rendir"
               End
               Begin VB.Menu opc11823 
                  Caption         =   "Vales Provisionales"
               End
            End
            Begin VB.Menu opc1183 
               Caption         =   "Operaciones Varias"
            End
            Begin VB.Menu opc1184 
               Caption         =   "Operaciones de Pago"
            End
            Begin VB.Menu opc1185 
               Caption         =   "Rendiciones"
            End
         End
         Begin VB.Menu opc11A 
            Caption         =   "Modifica recibos"
         End
      End
      Begin VB.Menu opc12 
         Caption         =   "Actualiza Tablas"
         Begin VB.Menu opc121 
            Caption         =   "Operaciones Generales"
         End
         Begin VB.Menu opc122 
            Caption         =   "Conceptos de Movimientos"
         End
         Begin VB.Menu opc123 
            Caption         =   "Codigo de Caja"
         End
         Begin VB.Menu opc124 
            Caption         =   " Bancos"
         End
         Begin VB.Menu opc125 
            Caption         =   "Cuentas por Bancos"
         End
         Begin VB.Menu opc126 
            Caption         =   "Empresas"
         End
         Begin VB.Menu opc127 
            Caption         =   "Grupo de Gastos"
            Visible         =   0   'False
         End
         Begin VB.Menu opc128 
            Caption         =   "Forma de Pago"
         End
      End
      Begin VB.Menu opc13 
         Caption         =   "Actualiza Saldos"
         Begin VB.Menu opc131 
            Caption         =   "Saldos Iniciales de Bancos"
         End
         Begin VB.Menu opc132 
            Caption         =   "Saldos Iniciales de Caja"
         End
      End
   End
   Begin VB.Menu opc2 
      Caption         =   "&Consultas"
      Begin VB.Menu opc21 
         Caption         =   "Recibos"
      End
      Begin VB.Menu opc22 
         Caption         =   "Rendiciones"
         Visible         =   0   'False
      End
      Begin VB.Menu opc23 
         Caption         =   "Busqueda General"
      End
   End
   Begin VB.Menu opc3 
      Caption         =   "&Reportes"
      Begin VB.Menu opc31 
         Caption         =   "Movimientos de Caja"
         Begin VB.Menu mnumovcajaconce 
            Caption         =   "Movimientos de Caja por Conceptos"
         End
         Begin VB.Menu mnumovdiari 
            Caption         =   "Movimientos Diarios"
         End
         Begin VB.Menu mnuChequesEmitidos 
            Caption         =   "Relacion de Cheques "
         End
         Begin VB.Menu opc32 
            Caption         =   "Resumen de Caja"
            Enabled         =   0   'False
            Visible         =   0   'False
         End
         Begin VB.Menu opc315 
            Caption         =   "Documentos x Rendir"
         End
      End
      Begin VB.Menu opc33 
         Caption         =   "Cuenta Corriente"
      End
      Begin VB.Menu opc34 
         Caption         =   "Impresion Recibos"
         Begin VB.Menu opc341 
            Caption         =   "Ingreso/Egreso"
         End
         Begin VB.Menu opc342 
            Caption         =   "Transferencia"
         End
      End
      Begin VB.Menu opc35 
         Caption         =   "Saldos x Proveedor y Cliente"
         Begin VB.Menu opc351 
            Caption         =   "Saldos x Proveedor"
         End
         Begin VB.Menu opc352 
            Caption         =   "Saldos x Cliente"
         End
      End
      Begin VB.Menu opc36 
         Caption         =   "Gastos "
      End
      Begin VB.Menu opc37 
         Caption         =   "Comprobantes"
         Begin VB.Menu opc371 
            Caption         =   "Comprobantes de Retencion"
         End
         Begin VB.Menu opc372 
            Caption         =   "Detraccion"
         End
      End
      Begin VB.Menu opc38 
         Caption         =   "Resumenes de Tesoreria"
      End
      Begin VB.Menu opc39 
         Caption         =   "Cuentas por rendir"
      End
   End
   Begin VB.Menu opc4 
      Caption         =   "&Procesos"
      Begin VB.Menu opc41 
         Caption         =   "Regeneracion de Saldos"
      End
      Begin VB.Menu opc42 
         Caption         =   "Cierre Mensual"
      End
      Begin VB.Menu opc43 
         Caption         =   "Proceso Cierre Diario"
      End
      Begin VB.Menu opc44 
         Caption         =   "Transferencia Contabilidad"
      End
      Begin VB.Menu opc45 
         Caption         =   "Telecredito"
         Visible         =   0   'False
         Begin VB.Menu opc451 
            Caption         =   "Generacion"
         End
         Begin VB.Menu opc452 
            Caption         =   "Actualizacion"
         End
         Begin VB.Menu opc453 
            Caption         =   "Reportes"
         End
      End
   End
   Begin VB.Menu opc5 
      Caption         =   "&Configuracion"
      Begin VB.Menu opc5_01 
         Caption         =   "Configurar empresa"
      End
   End
   Begin VB.Menu opc6 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "MDIPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Bar1_ButtonClick(ByVal Button As MSComctlLib.Button)
   Select Case Button.Index
      Case 1
        'Call opc12_Click
      Case 2
        'Call opc11_Click
      Case 3
        'Call opc13_Click
      Case 4
      Case 5
      Case 6
         Call opc6_Click
   End Select
End Sub
Private Sub mnu00_01_01_Click(Index As Integer)
    Call Screen.ActiveForm.Pavant(Index)
End Sub
Private Sub mnu00_01_Click(Index As Integer)
    Call Screen.ActiveForm.PMant(Index)
End Sub
Private Sub MDIForm_Load()
   Unload FrmIngreso
'   MostrarForm Me, "M"
End Sub
Private Sub MDIForm_Unload(Cancel As Integer)
    Set cbdatos = Nothing
    Set VGCNx = Nothing
    Set VGGeneral = Nothing
    Set VGCnxCT = Nothing
    End
End Sub
Private Sub mnudocanula_Click()
    FrmRepDocAnula.Show
End Sub
Private Sub mnudocumentos_Click()
    FrmDocumento.Show
End Sub
Private Sub mnutipcambio_Click()
    frmMantTipoCambio.Show
End Sub
Private Sub mnuChequesEmitidos_Click()
   frmRepChequesEmitidos.Show
End Sub
Private Sub mnumovcajaconce_Click()
    FrmRepCajaBancos.Show
End Sub
Private Sub opc112_Click()
   FrmMovimientoClientes.Show
End Sub
Private Sub opc113_Click()
   FrmMovimientoProve.Show
End Sub

Private Sub mnumovdiari_Click()
FrmLiquidarionDiaria.resumido = 0
FrmLiquidarionDiaria.Show
End Sub

Private Sub opc1111_Click()
FrmMovimientoCaja.fondofijo = 0
FrmMovimientoCaja.docxrendir = 0
FrmMovimientoCaja.Show
End Sub
Private Sub opc1112_Click()
FrmMovimientoClientes.Show
End Sub
Private Sub opc1113_Click()
FrmMovimientoProve.fondofijo = 0
FrmMovimientoProve.docxrendir = 0
FrmMovimientoProve.Show
End Sub
Private Sub opc1141_Click()
   FrmConciliacionBancos.Show
End Sub
Private Sub opc1142_Click()
   FrmRendicionCaja.Show
End Sub
Private Sub opc1151_Click()
   frmTransferencias.CasoOrigen = "B"
   frmTransferencias.CasoDestino = "B"
   frmTransferencias.Show
End Sub

Private Sub opc1152_Click()
   frmTransferencias.CasoOrigen = "C"
   frmTransferencias.CasoDestino = "B"
   frmTransferencias.cuentasxrendir = 0
   frmTransferencias.fondofijo = 0
   frmTransferencias.tipo = 0
   frmTransferencias.Show
End Sub

Private Sub opc1153_Click()
   frmTransferencias.CasoOrigen = "B"
   frmTransferencias.CasoDestino = "C"
   frmTransferencias.cuentasxrendir = 0
   frmTransferencias.fondofijo = 0
   frmTransferencias.tipo = 0
   frmTransferencias.Show
End Sub

Private Sub opc1154_Click()
   frmTransferencias.CasoOrigen = "C"
   frmTransferencias.CasoDestino = "C"
   frmTransferencias.cuentasxrendir = 0
   frmTransferencias.fondofijo = 0
   frmTransferencias.tipo = 0
   frmTransferencias.titulo = " Transferencia de Caja Principal a Otra Caja "
   frmTransferencias.Show
End Sub
Private Sub opc1161_Click()
  frmAnularBorraRecibos.Show
End Sub
Private Sub opc1163_Click()
frmAnularTransferencia.Show
End Sub

Private Sub opc11711_Click()
   frmTransferencias.CasoOrigen = "B"
   frmTransferencias.CasoDestino = "C"
   frmTransferencias.cuentasxrendir = 1
   frmTransferencias.fondofijo = 0
   frmTransferencias.tipo = 0
   frmTransferencias.Show
End Sub

Private Sub opc11712_Click()
   frmTransferencias.CasoOrigen = "C"
   frmTransferencias.CasoDestino = "C"
   frmTransferencias.cuentasxrendir = 1
   frmTransferencias.fondofijo = 0
   frmTransferencias.tipo = 0
   frmTransferencias.Show

End Sub

Private Sub opc11721_Click()
   frmTransferencias.CasoOrigen = "C"
   frmTransferencias.CasoDestino = "B"
   frmTransferencias.cuentasxrendir = 1
   frmTransferencias.fondofijo = 0
   frmTransferencias.tipo = 1
   frmTransferencias.Show
End Sub

Private Sub opc11722_Click()
   frmTransferencias.CasoOrigen = "C"
   frmTransferencias.CasoDestino = "C"
   frmTransferencias.cuentasxrendir = 1
   frmTransferencias.fondofijo = 0
   frmTransferencias.tipo = 1
   frmTransferencias.Show
End Sub



Private Sub opc1173_Click()
FrmMovimientoCaja.fondofijo = 0
FrmMovimientoCaja.docxrendir = 1
FrmMovimientoCaja.Show
End Sub

Private Sub opc1174_Click()
  frmMantprovision.cuentasxrendir = 1
  frmMantprovision.fondofijo = 0
  frmMantprovision.Show
End Sub

Private Sub opc1175_Click()
FrmMovimientoProve.fondofijo = 0
FrmMovimientoProve.docxrendir = 1
FrmMovimientoProve.Show
End Sub

Private Sub opc11811_Click()
   frmTransferencias.CasoOrigen = "B"
   frmTransferencias.CasoDestino = "C"
   frmTransferencias.cuentasxrendir = 0
   frmTransferencias.fondofijo = 1
   frmTransferencias.tipo = 0
   frmTransferencias.Show

End Sub

Private Sub opc11812_Click()
   frmTransferencias.CasoOrigen = "C"
   frmTransferencias.CasoDestino = "C"
   frmTransferencias.cuentasxrendir = 0
   frmTransferencias.fondofijo = 1
   frmTransferencias.tipo = 0
   frmTransferencias.Show
End Sub

Private Sub opc11813_Click()
   frmTransferencias.CasoOrigen = "C"
   frmTransferencias.CasoDestino = "C"
   frmTransferencias.cuentasxrendir = 0
   frmTransferencias.fondofijo = 1
   frmTransferencias.tipo = 2
   frmTransferencias.Show
End Sub

Private Sub opc11821_Click()
   frmTransferencias.CasoOrigen = "B"
   frmTransferencias.CasoDestino = "C"
   frmTransferencias.cuentasxrendir = 1
   frmTransferencias.fondofijo = 0
   frmTransferencias.tipo = 1
   frmTransferencias.Show

End Sub

Private Sub opc11822_Click()
   frmTransferencias.CasoOrigen = "C"
   frmTransferencias.CasoDestino = "C"
   frmTransferencias.cuentasxrendir = 1
   frmTransferencias.fondofijo = 0
   frmTransferencias.tipo = 1
   frmTransferencias.Show
End Sub
Private Sub opc11823_Click()
   frmTransferencias.CasoOrigen = "C"
   frmTransferencias.CasoDestino = "C"
   frmTransferencias.cuentasxrendir = 0
   frmTransferencias.fondofijo = 1
   frmTransferencias.tipo = 2
   frmTransferencias.Show
End Sub

Private Sub opc1183_Click()
FrmMovimientoCaja.fondofijo = 1
FrmMovimientoCaja.docxrendir = 0
FrmMovimientoCaja.Show
End Sub

Private Sub opc1184_Click()
 frmMantprovision.fondofijo = 1
 frmMantprovision.cuentasxrendir = 0
 frmMantprovision.Show
End Sub

Private Sub opc11A_Click()
 Screen.MousePointer = vbHourglass
  frmModrecibos.Show
    Screen.MousePointer = vbDefault
End Sub

Private Sub opc121_Click()
  FrmOperacion.Show
End Sub
Private Sub opc122_Click()
  FrmConceptocaja.Show
End Sub
Private Sub opc123_Click()
  FrmCodigocajas.Show
End Sub
Private Sub opc124_Click()
  frmBanco.Show
End Sub
Private Sub opc125_Click()
   FrmCuentaBancaria.Show
End Sub

Private Sub opc126_Click()
  FrmEmpresa.Show
End Sub

Private Sub opc128_Click()
FrmFormadePago.Show
End Sub


Private Sub opc14_Click()
  FrmCopiaPedido.Show
End Sub

Private Sub opc131_Click()
   Frmsaldoinicial.CmbOper.ListIndex = 1
   Frmsaldoinicial.Show
End Sub

Private Sub opc132_Click()
   Frmsaldoinicial.CmbOper.ListIndex = 0
   Frmsaldoinicial.Show
End Sub

Private Sub opc21_Click()
  Screen.MousePointer = vbHourglass
    frmMantRecibos.Show
    Screen.MousePointer = vbDefault
End Sub

Private Sub opc22_Click()
  FrmRepRendiciones.Show
  End Sub

Private Sub opc23_Click()
    FrmBusqueda.Show
    End Sub

Private Sub opc315_Click()
FrmDocxrendir.Show
End Sub

Private Sub opc32_Click()
  CstVentas.Show
End Sub

Private Sub opc33_Click()
  frmRepCtaCteCajaBancos.Show
End Sub

Private Sub opc341_Click()
  FrmImprimirRecibo.Show
End Sub

Private Sub opc342_Click()
  frmRepTransferencia.Show
End Sub

Private Sub opc411_Click()
    FrmRepVtasxArt.Show
End Sub

Private Sub opc412_Click()
    FrmRepVtasxFact.Show
End Sub

Private Sub opc413_Click()
    FrmRepGuiaFactBol.Show
End Sub

Private Sub opc351_Click()
   RptSaldoxProveedor.Show
End Sub

Private Sub opc352_Click()
   RptSaldoxCliente.Show
End Sub

Private Sub opc36_Click()
Frmreplistgastos.Show 1
End Sub
Private Sub opc371_Click()
frmRepComprobantesRetencion.Show
End Sub

Private Sub opc38_Click()
FrmResumenGastosTesoreria.Show
End Sub

Private Sub opc39_Click()
frmCtasxrendir.Show 1
End Sub

Private Sub opc42_Click()
 FrmCierreMensual.Show 1
End Sub
Private Sub opc43_Click()
   FrmCierreDiario.Show 1
End Sub
Private Sub opc511_Click()
   FrmEmpresa.Show 1
End Sub
Private Sub opc512_Click()
   FrmAlmacen.Show 1
End Sub
Private Sub opc513_Click()
   frmBanco.Show 1
End Sub

Private Sub opc515_Click()
   FrmFormaPago.Show 1
End Sub

Private Sub opc516_Click()
   FrmMoneda.Show
End Sub

Private Sub opc517_Click()
 FrmNegocio.Show 1
End Sub

Private Sub opc518_Click()
  FrmPuntoVenta.Show
End Sub

Private Sub opc519_Click()
  FrmUnidadMedida.Show
End Sub

Private Sub opc51A_Click()
 FrmVendedor.Show
End Sub

Private Sub opc51B_Click()
 FrmZona.Show
End Sub

Private Sub opc51C_Click()
  FrmUsuario.Show
End Sub

Private Sub opc521_Click()
 FrmProducto.Show
End Sub

Private Sub opc522_Click()
 Frmcliente.Show
End Sub

Private Sub opc523_Click()
 FrmModoVenta.Show
End Sub

Private Sub opc524_Click()
 FrmParametroVenta.Show
End Sub

Private Sub opc525_Click()
 FrmLimiteCredito.Show
End Sub
Private Sub opc531_Click()
 FrmPtoVtaDoc.Show
End Sub
Private Sub opc532_Click()
 FrmSerieDocumento.Show
End Sub
Private Sub opc533_Click()
 FrmZonaVendedor.Show
End Sub
Private Sub opc44_Click()
 FrmContabilizaTesoreria.Show 1
End Sub
Private Sub opc451_Click()
VGmodifica = 0
FrmPagosinternet.Show 1
End Sub
Private Sub opc452_Click()
VGmodifica = 1
FrmPagosinternet.Show 1
End Sub
Private Sub opc453_Click()
FrmTelecreditoreportes.Show 1
End Sub
Private Sub opc5_01_Click()
VGtipo = caja
FrmCfgEmpresa.Show
End Sub
Private Sub opc6_Click()
   If MsgBox("Desea Salir del Sistema?", vbYesNo, "AVISO") = vbYes Then
      Set cbdatos = Nothing
      Set VGCNx = Nothing
      Set VGGeneral = Nothing
      Set VGCnxCT = Nothing
      End
   End If
End Sub

Private Sub opc7111_Click()
   frmTransferencias.CasoOrigen = "B"
   frmTransferencias.CasoDestino = "C"
   frmTransferencias.cuentasxrendir = 1
   frmTransferencias.Show
End Sub

Private Sub opc7112_Click()
   frmTransferencias.CasoOrigen = "C"
   frmTransferencias.CasoDestino = "C"
   frmTransferencias.cuentasxrendir = 1
   frmTransferencias.Show
End Sub

Private Sub Panel_PanelClick(ByVal Panel As MSComctlLib.Panel)
Select Case Panel.Index
        Case 1, 2
            Frmseleccionfechatrabajo.Show 1
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


