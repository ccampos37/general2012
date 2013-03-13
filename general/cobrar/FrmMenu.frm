VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form MDIPrincipal1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   Caption         =   "Sistema de Cuentas Por Cobrar"
   ClientHeight    =   7800
   ClientLeft      =   225
   ClientTop       =   825
   ClientWidth     =   11355
   ControlBox      =   0   'False
   Icon            =   "FrmMenu.frx":0000
   LinkTopic       =   "MDIForm1"
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   7800
   ScaleMode       =   0  'User
   ScaleWidth      =   11355
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11355
      _ExtentX        =   20029
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      Wrappable       =   0   'False
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   3
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin Crystal.CrystalReport CryRptProc 
      Left            =   4920
      Top             =   3480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin ComctlLib.StatusBar Panel 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   1
      Top             =   7485
      Width           =   11355
      _ExtentX        =   20029
      _ExtentY        =   556
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   5
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   6174
            MinWidth        =   6174
            Text            =   "Empresa : CAMTEX S.A."
            TextSave        =   "Empresa : CAMTEX S.A."
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   5292
            MinWidth        =   5292
            Text            =   "Usuario : Administrador"
            TextSave        =   "Usuario : Administrador"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   5292
            MinWidth        =   5292
            Text            =   "Base : DESARROLLO"
            TextSave        =   "Base : DESARROLLO"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   5292
            MinWidth        =   5292
            Text            =   "Fecha : 23/09/2002 10:00:03 am."
            TextSave        =   "Fecha : 23/09/2002 10:00:03 am."
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMenu.frx":030A
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Crystal.CrystalReport oCrystalReport 
      Left            =   4170
      Top             =   7290
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   5250
      Top             =   7260
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   3
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMenu.frx":0624
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMenu.frx":093E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMenu.frx":0C58
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu Opc1 
      Caption         =   "Movimientos"
      Begin VB.Menu Opc11 
         Caption         =   "Ingreso Datos"
         Begin VB.Menu Opc111 
            Caption         =   "Planilla Cobranzas"
            Begin VB.Menu Opc1111 
               Caption         =   "Ingreso Documentos"
            End
            Begin VB.Menu Opc1112 
               Caption         =   "Elimina Documentos de Planilla"
            End
         End
         Begin VB.Menu Opc112 
            Caption         =   "Documentos Varios"
            Begin VB.Menu Opc1121 
               Caption         =   "Ingreso Documentos"
            End
            Begin VB.Menu Opc1122 
               Caption         =   "Elimina Documentos de Planilla"
            End
         End
         Begin VB.Menu opt1 
            Caption         =   "-"
         End
         Begin VB.Menu Opc113 
            Caption         =   "Nota Abono/Cargo"
            Begin VB.Menu Opc1131 
               Caption         =   "Ingresa Documento en Cta. Cte."
            End
            Begin VB.Menu Opc1132 
               Caption         =   "Anula Documento Registrado"
            End
            Begin VB.Menu Opc1133 
               Caption         =   "Elimina Documento Registrado"
            End
         End
         Begin VB.Menu Opc115 
            Caption         =   "Nota Abono/Cargo Fisico"
         End
         Begin VB.Menu Opc114 
            Caption         =   "Canje Renovacion"
            Begin VB.Menu Opc1141 
               Caption         =   "Canje de Documentos"
            End
            Begin VB.Menu Opc1142 
               Caption         =   "Renovacion Documentos"
            End
         End
      End
      Begin VB.Menu Opc12 
         Caption         =   "Actualiza Tablas"
         Begin VB.Menu Opc121 
            Caption         =   "Tabla Bancos"
         End
         Begin VB.Menu Opc122 
            Caption         =   "Tabla Tipos Documentos"
         End
         Begin VB.Menu Opc123 
            Caption         =   "Tabla de Conceptos"
            Visible         =   0   'False
         End
         Begin VB.Menu Opc124 
            Caption         =   "Tabla de Vendedores"
         End
         Begin VB.Menu Opc125 
            Caption         =   "Tabla de Empresas"
         End
         Begin VB.Menu Opc126 
            Caption         =   "Tabla de Zonas"
            Visible         =   0   'False
         End
         Begin VB.Menu Opc127 
            Caption         =   "Tabla de Tipo de Negocio"
         End
         Begin VB.Menu Opc128 
            Caption         =   "Tabla Tipo Planillas"
         End
         Begin VB.Menu mnulimicred 
            Caption         =   "Tablas Limite de Credito"
            Begin VB.Menu mnugruplimicred 
               Caption         =   "Grupo limite Credito"
            End
            Begin VB.Menu mnuDocxgrupcred 
               Caption         =   "Documento x Grupo de Credito"
            End
         End
      End
      Begin VB.Menu Opc13 
         Caption         =   "Actualiza Maestros"
         Begin VB.Menu Opc131 
            Caption         =   "Clientes"
         End
         Begin VB.Menu Opc132 
            Caption         =   "Limite Credito"
         End
         Begin VB.Menu Opc133 
            Caption         =   "Direcciones Clientes"
         End
         Begin VB.Menu Opc134 
            Caption         =   "Cliente Grupo Limite Credito"
         End
      End
   End
   Begin VB.Menu opc2 
      Caption         =   "Procesos"
      Begin VB.Menu opc21 
         Caption         =   "Cierre Mensual"
         Enabled         =   0   'False
      End
      Begin VB.Menu opc22 
         Caption         =   "Regularizacion Facturas"
         Enabled         =   0   'False
      End
      Begin VB.Menu opc23 
         Caption         =   "Regeneracion Saldos"
      End
      Begin VB.Menu opc24 
         Caption         =   "Anulacion de Letras"
         Visible         =   0   'False
      End
      Begin VB.Menu opc25 
         Caption         =   "Actualiza Tipo Cambio"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu opc3 
      Caption         =   "Reportes"
      Begin VB.Menu opc31 
         Caption         =   "Saldo Documentos"
         Begin VB.Menu opc311 
            Caption         =   "Saldo por Cliente"
         End
         Begin VB.Menu opc312 
            Caption         =   "Saldo por Vendedor"
         End
      End
      Begin VB.Menu opc32 
         Caption         =   "Estado Cta Cte"
         Begin VB.Menu opc321 
            Caption         =   "Cta Cte x Vendedor"
            Enabled         =   0   'False
            Visible         =   0   'False
         End
         Begin VB.Menu opc322 
            Caption         =   "Cta Cte x Clientes"
         End
      End
      Begin VB.Menu opc33 
         Caption         =   "Planilla Cobranza"
      End
      Begin VB.Menu menu03_11 
         Caption         =   "Planilla de Cobranza - Bancos"
      End
      Begin VB.Menu opc34 
         Caption         =   "Planilla Varios"
      End
      Begin VB.Menu opc3B 
         Caption         =   "Modo Venta"
      End
      Begin VB.Menu opc35 
         Caption         =   "Resumen Planilla Cobranza"
         Enabled         =   0   'False
         Visible         =   0   'False
         Begin VB.Menu opc351 
            Caption         =   "Resumen Diario de Cobranzas"
         End
         Begin VB.Menu opc352 
            Caption         =   "Resumen Detallado de Cobranza"
         End
      End
      Begin VB.Menu opc36 
         Caption         =   "Documentos"
         Visible         =   0   'False
         Begin VB.Menu mnuavicobra 
            Caption         =   "Aviso de Cobranzas"
         End
         Begin VB.Menu opc361 
            Caption         =   "Listado General"
         End
         Begin VB.Menu opc362 
            Caption         =   "Vencidos/Por Vencer"
         End
         Begin VB.Menu opc364 
            Caption         =   "Vencidos x Vencer (Otro Formato)"
            Enabled         =   0   'False
         End
         Begin VB.Menu opc363 
            Caption         =   "Por Aplicar"
            Visible         =   0   'False
         End
      End
      Begin VB.Menu opc3D 
         Caption         =   "Resumen de Deudas x Cliente"
      End
      Begin VB.Menu opc37 
         Caption         =   "Nota Abono/Cargo"
         Enabled         =   0   'False
      End
      Begin VB.Menu opc38 
         Caption         =   "Clientes Reportes"
      End
      Begin VB.Menu opc39 
         Caption         =   "Planilla de Canjes"
         Begin VB.Menu opc391 
            Caption         =   "Planilla de Canjes"
         End
         Begin VB.Menu opc932 
            Caption         =   "Documentos Canjeados"
         End
      End
      Begin VB.Menu opc3A 
         Caption         =   "Planilla de Renovacion"
      End
      Begin VB.Menu opc3C 
         Caption         =   "Letras"
         Visible         =   0   'False
         Begin VB.Menu opc3C1 
            Caption         =   "Letras Descontadas"
            Enabled         =   0   'False
         End
         Begin VB.Menu opc3C2 
            Caption         =   "Impresión de Letras"
         End
      End
   End
   Begin VB.Menu opc4 
      Caption         =   "Consultas"
      Begin VB.Menu opc41 
         Caption         =   "Saldo por Cliente"
      End
   End
   Begin VB.Menu opc5 
      Caption         =   "Salida"
   End
End
Attribute VB_Name = "MDIPrincipal1"
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
    frmrepavisocobra.Show
End Sub

Private Sub mnuDocxgrupcred_Click()
   frmDocxLimit.Show
End Sub

Private Sub mnugruplimicred_Click()
   fmrlimitgrupo.Show
End Sub

Private Sub Opc1111_Click()
  FrmPlanillaCobranza.Show
End Sub

Private Sub Opc1112_Click()
  FrmPlanillaCobranzaModi.Show
End Sub

Private Sub Opc1121_Click()
   FrmPlanillaVarios.Show
End Sub

Private Sub Opc1122_Click()
  FrmPlanillaVariosModi.Show
End Sub

Private Sub Opc1131_Click()
 FrmNotas.Show
End Sub

Private Sub Opc1132_Click()
  FrmAnulaNota.Show
End Sub

Private Sub Opc1133_Click()
  FrmEliminaNota.Show
End Sub

Private Sub Opc1134_Click()
   
End Sub

Private Sub Opc1141_Click()
  FrmPlanillaCanjes.Show
End Sub

Private Sub Opc1142_Click()
  FrmPlanillaRenova.Show
End Sub

Private Sub Opc115_Click()
  FrmNotaFisico.Show
End Sub

Private Sub Opc121_Click()
  frmBanco.Show
End Sub

Private Sub Opc122_Click()
  FrmTipodocumentos.Show
End Sub

Private Sub Opc123_Click()
  FrmTipoConcepto.Show
End Sub

Private Sub Opc124_Click()
  FrmVendedor.Show
End Sub

Private Sub Opc125_Click()
  FrmEmpresa.Show
End Sub

Private Sub Opc126_Click()
  FrmZona.Show
End Sub

Private Sub Opc127_Click()
 FrmNegocio.Show
End Sub

Private Sub Opc128_Click()
  FrmTipoPlanilla.Show
End Sub

Private Sub Opc131_Click()
 Frmcliente.Show
End Sub

Private Sub Opc132_Click()
  FrmLimiteCredito.Show
End Sub

Private Sub Opc133_Click()
 FrmMultidireccion.Show
End Sub

Private Sub Opc134_Click()
   frmClientexGrupoCred.Show
End Sub

Private Sub opc23_Click()
  If MsgBox("Desea Regenerar los Saldos?", vbYesNo, MsgTitle) = vbYes Then
     PrcGeneraSaldos.Show 1
  End If
End Sub

Private Sub opc24_Click()
   frmAnularLetras.Show
End Sub

Private Sub opc25_Click()
 Dim SQL As String
   Screen.MousePointer = 11
   SQL = "insert ct_tipocambio "
   SQL = SQL & "select * from " & g_BaseContab & ".dbo.ct_tipocambio where tipocambiofecha not in"
   SQL = SQL & "(select tipocambiofecha from ct_tipocambio)"
   VGCNx.Execute (SQL)
   Screen.MousePointer = 1
End Sub

Private Sub opc311_Click()
  RptSaldoxCliente.Show
End Sub

Private Sub opc312_Click()
  RptSaldoxVendedor.Show
End Sub

Private Sub opc321_Click()
    RptCtactexVendedor.Show
End Sub

Private Sub opc322_Click()
    RptctactexCliente.Show
End Sub

Private Sub opc33_Click()
  FrmRepPlanillaCob.Show
End Sub

Private Sub opc34_Click()
    FrmRepPlanillaDocVar.Show
End Sub

Private Sub opc351_Click()
    RptResumenCobranzaDiaria.Show
End Sub

Private Sub opc352_Click()
    RptResumenCobranzaDetallada.Show
End Sub

Private Sub opc361_Click()
  frmRepListadoDocumentos.Show
End Sub

Private Sub opc362_Click()
  RptDocumentosxCobrar.Show
End Sub

Private Sub opc363_Click()
  RptDocumentosxAplicar.Show
End Sub

Private Sub opc364_Click()
 FrmRepDocvenciXvence.Show
End Sub

Private Sub opc37_Click()
  RptNotaabono.Show
End Sub

Private Sub opc382_Click()
    Rptclientexzona.Show
End Sub

Private Sub opc383_Click()
   RptclientexVendedor.Show
End Sub

Private Sub opc384_Click()
   RptClientexdistrito.Show
End Sub

Private Sub opc385_Click()
  Rptclientexcategoria.Show
End Sub

Private Sub opc38_Click()
   frmRepClientes.Show
End Sub

Private Sub opc391_Click()
  FrmRepPlanillaCanjeRenovacion.Opcion = "1"
  FrmRepPlanillaCanjeRenovacion.Show
End Sub


Private Sub opc3D_Click()
   frmrepantigdeudas.Show
End Sub

Private Sub opc932_Click()
  FrmRepOtroPlanillaCanjeRenovacion.Opcion = "1"
  FrmRepOtroPlanillaCanjeRenovacion.Show
End Sub

Private Sub opc3A_Click()
  FrmRepOtroPlanillaCanjeRenovacion.Opcion = "2"
  FrmRepOtroPlanillaCanjeRenovacion.Show
End Sub

Private Sub opc3B_Click()
  RptModoVenta.Show
End Sub

Private Sub opc3C1_Click()
  ' frmRepLetrasDescontadas.Show
End Sub

Private Sub opc3C2_Click()
  frmRepImpresionLetras.Show
End Sub

Private Sub opc41_Click()
  CstSaldoCliente.Show
End Sub

Private Sub opc5_Click()
   If MsgBox("Desea Salir del Sistema?", vbYesNo, "AVISO") = vbYes Then
      Set cn = Nothing
      Set VGGeneral = Nothing
      Set VGCNx = Nothing
      Set VGcnxCT = Nothing
      End
   End If
End Sub

Private Sub Panel_PanelClick(ByVal Panel As ComctlLib.Panel)
  If Panel.Index = 5 Then
     Load FrmIngreso
     FrmIngreso.Show 1
  End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
   Select Case Button.Index
      Case 1
        Call Opc1131_Click
      Case 2
        'Call opc11_Click
      Case 3
        Call opc5_Click
   End Select
End Sub
