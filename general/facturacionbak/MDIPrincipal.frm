VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.MDIForm MDIPrincipal 
   BackColor       =   &H00C0FFC0&
   Caption         =   "Sistema de Facturacion"
   ClientHeight    =   7455
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   12615
   Icon            =   "MDIPrincipal.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImgMenu 
      Left            =   1125
      Top             =   3105
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":0CCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":110A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":170B
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":1F5C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":2843
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":306B
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Bar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "ImgMenu"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Facturacion"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Clientes"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Registro de Ventas"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Reporte Gerencial"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Lista Precios"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Documentos"
            ImageIndex      =   6
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   405
      Top             =   4590
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":38DD
            Key             =   "Facturar"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":CE75
            Key             =   "Modificar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":1C2CE
            Key             =   "Nuevo"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":26F2F
            Key             =   "Eliminar"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":3A243
            Key             =   "Facturado"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":44B80
            Key             =   "Retornar"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":572D5
            Key             =   "Almacen"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":6D09F
            Key             =   "Imprimir"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":7FC59
            Key             =   "Buscar"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":A2DB3
            Key             =   "Copia"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":AD84D
            Key             =   "Usuarios"
         EndProperty
      EndProperty
   End
   Begin Crystal.CrystalReport CryRptProc 
      Left            =   5160
      Top             =   2160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport oCrystalReport 
      Left            =   1080
      Top             =   8040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   7140
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7056
            MinWidth        =   7056
            Text            =   "Mes :"
            TextSave        =   "Mes :"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
            Text            =   "Año :"
            TextSave        =   "Año :"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
            Text            =   "fecha de Trabajo"
            TextSave        =   "fecha de Trabajo"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
            Text            =   "Servidor"
            TextSave        =   "Servidor"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   5292
            MinWidth        =   5292
            Text            =   "Base de datos"
            TextSave        =   "Base de datos"
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   10710
      Top             =   8190
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":C0556
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":C09AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":C0DFE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList3 
      Left            =   1215
      Top             =   4635
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":C1256
            Key             =   "Buscar"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":C1B76
            Key             =   "Copiar"
         EndProperty
      EndProperty
   End
   Begin VB.Menu opc1 
      Caption         =   "Movimientos"
      Begin VB.Menu opc1_01 
         Caption         =   "Lista Precios"
      End
      Begin VB.Menu opc1_02 
         Caption         =   "Pedido"
         Visible         =   0   'False
      End
      Begin VB.Menu opc1_07 
         Caption         =   "Pedido Ventanilla"
      End
      Begin VB.Menu opc1_08 
         Caption         =   "Modifica Pedido"
      End
      Begin VB.Menu opc1_03 
         Caption         =   "Anula Factura"
      End
      Begin VB.Menu opc1_04 
         Caption         =   "Copia Documentos"
      End
      Begin VB.Menu opc1_05 
         Caption         =   "Impresion de Guias"
         Visible         =   0   'False
      End
      Begin VB.Menu Opc1_06 
         Caption         =   "Generacion de Pedidos x Liq. Compras"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu opc4 
      Caption         =   "Reportes"
      Begin VB.Menu opc4_01 
         Caption         =   "&Movimientos de Ventas"
         Begin VB.Menu opc4_01_01 
            Caption         =   "&Ventas por Articulo"
         End
         Begin VB.Menu opc4_01_09 
            Caption         =   "Ventas Cliente x Articulo"
            Enabled         =   0   'False
            Visible         =   0   'False
         End
         Begin VB.Menu opc4_01_02 
            Caption         =   "&Ventas por Factura"
            Enabled         =   0   'False
            Visible         =   0   'False
         End
         Begin VB.Menu opc4_01_12 
            Caption         =   "Ventas Factura Detallado"
            Visible         =   0   'False
         End
         Begin VB.Menu opc4_01_03 
            Caption         =   "&Guias/Facturas/Boletas"
            Visible         =   0   'False
         End
         Begin VB.Menu opc4_01_04 
            Caption         =   "Comisisones Vendedores"
         End
         Begin VB.Menu opc4_01_11 
            Caption         =   "Numeracion por Documentos"
         End
         Begin VB.Menu opc4_01_10 
            Caption         =   "Documentos Anulados"
         End
         Begin VB.Menu opc4_01_08 
            Caption         =   "&Ventas por Forma de Pago"
         End
         Begin VB.Menu mnuprecio 
            Caption         =   "Variacion &Precio"
            Enabled         =   0   'False
            Visible         =   0   'False
         End
      End
      Begin VB.Menu opc4_02 
         Caption         =   "Estadisticas Mensuales"
         Visible         =   0   'False
         Begin VB.Menu opc2_3_1 
            Caption         =   "Clientes x Tipo de Negocio"
         End
         Begin VB.Menu Prod_negocio 
            Caption         =   "Productos x tipo de Negocio"
         End
         Begin VB.Menu Prod_clientes 
            Caption         =   "Productos x Clientes"
         End
      End
      Begin VB.Menu opc4_10 
         Caption         =   "Estadisticas"
         Begin VB.Menu opc4_10_01 
            Caption         =   "&Clientes"
         End
         Begin VB.Menu opc4_10_02 
            Caption         =   "&Articulos"
         End
         Begin VB.Menu opc4_10_03 
            Caption         =   "&Vendedores"
         End
         Begin VB.Menu opc4_10_04 
            Caption         =   "&Negocios"
            Visible         =   0   'False
         End
         Begin VB.Menu opc4_10_05 
            Caption         =   "Reporte Gerente"
         End
         Begin VB.Menu opc4_10_06 
            Caption         =   "&Resumen Ventas Mensuales"
         End
      End
      Begin VB.Menu opc4_06 
         Caption         =   "&Registro Ventas"
      End
      Begin VB.Menu opc4_04 
         Caption         =   "&Notas  de Creditos"
      End
      Begin VB.Menu mnucrepcaj 
         Caption         =   "Reporte de Caja"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnufv 
         Caption         =   "Formato de Venta"
         Visible         =   0   'False
      End
      Begin VB.Menu opc4_05 
         Caption         =   "Ingreso por Contacto"
      End
   End
   Begin VB.Menu opc3 
      Caption         =   "Consultas"
      Begin VB.Menu opc3_1 
         Caption         =   "&Documentos"
      End
      Begin VB.Menu opc3_2 
         Caption         =   "&Ventas"
      End
      Begin VB.Menu opc3_3 
         Caption         =   "&Limite Credito"
      End
      Begin VB.Menu opc3_4 
         Caption         =   "Disponibilidad de  Producto"
      End
      Begin VB.Menu opc3_5 
         Caption         =   "Analisis  de ventas"
      End
   End
   Begin VB.Menu opc02 
      Caption         =   "Procesos"
      Begin VB.Menu opc02_1 
         Caption         =   "&Correccion datos generales"
      End
      Begin VB.Menu opc02_2 
         Caption         =   "Co&tizacion Libre"
      End
      Begin VB.Menu opc02_3 
         Caption         =   "&Eliminar Documento"
      End
      Begin VB.Menu opc02_4 
         Caption         =   "&Traslado entre Almacen"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu opt02_7 
         Caption         =   "-"
      End
      Begin VB.Menu opc02_6 
         Caption         =   "&Exportar Data"
         Visible         =   0   'False
      End
      Begin VB.Menu opc02_5 
         Caption         =   "&Transferencia de Datos"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuhshshs 
         Caption         =   "Precios a maeart"
         Visible         =   0   'False
      End
      Begin VB.Menu opc02_8 
         Caption         =   "Correccion de Forma de Pago"
      End
   End
   Begin VB.Menu opc05 
      Caption         =   "Configuracion"
      Begin VB.Menu opc05_1 
         Caption         =   "Tablas Basicas"
         Begin VB.Menu opc05_1_1 
            Caption         =   "&Empresa"
         End
         Begin VB.Menu opc05_1_2 
            Caption         =   "&Almacen"
         End
         Begin VB.Menu opc05_1_3 
            Caption         =   "&Bancos"
         End
         Begin VB.Menu opc05_1_4 
            Caption         =   "&Grupo Venta"
         End
         Begin VB.Menu opc05_1_5 
            Caption         =   "&Forma Pago"
         End
         Begin VB.Menu opc05_1_6 
            Caption         =   "&Moneda"
         End
         Begin VB.Menu opc05_1_7 
            Caption         =   "&Negocio"
         End
         Begin VB.Menu opc05_1_8 
            Caption         =   "&Punto Venta"
         End
         Begin VB.Menu opc05_1_9 
            Caption         =   "&Unidad Medida"
         End
         Begin VB.Menu opc05_1_A 
            Caption         =   "&Vendedor"
         End
         Begin VB.Menu opc05_1_B 
            Caption         =   "&Zona"
         End
         Begin VB.Menu opc05_1_c 
            Caption         =   "&Documentos"
         End
         Begin VB.Menu opc05_1_d 
            Caption         =   "Conceptos de Pago"
         End
      End
      Begin VB.Menu opc05_2 
         Caption         =   "Maestros"
         Begin VB.Menu opc05_2_2 
            Caption         =   "Clientes"
         End
         Begin VB.Menu opc05_2_3 
            Caption         =   "Modo Venta"
         End
         Begin VB.Menu opc05_2_4 
            Caption         =   "Parametros Venta"
         End
         Begin VB.Menu opc05_2_5 
            Caption         =   "Limites Creditos"
         End
      End
      Begin VB.Menu opc05_3 
         Caption         =   "Tablas Relaciones"
         Begin VB.Menu opc05_3_1 
            Caption         =   "Punto Venta-Documento"
         End
         Begin VB.Menu opc05_3_2 
            Caption         =   "Serie-Documentos"
         End
         Begin VB.Menu opc05_3_3 
            Caption         =   "Zona-Vendedor"
         End
         Begin VB.Menu opc05_3_4 
            Caption         =   "Ventas x Zona"
         End
      End
      Begin VB.Menu opc05_5 
         Caption         =   "Configuracion Empresa"
      End
      Begin VB.Menu opc05_6 
         Caption         =   "Configuracion Usuarios"
      End
      Begin VB.Menu opc05_7 
         Caption         =   "Configurar impresora"
      End
   End
   Begin VB.Menu opc6 
      Caption         =   "Salir"
   End
   Begin VB.Menu opc_priebas 
      Caption         =   "pruebas"
      Enabled         =   0   'False
      Visible         =   0   'False
   End
End
Attribute VB_Name = "MDIPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tipocorreccion As Integer
Private Sub Bar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
    Case 1
        Call opc1_02_Click
    Case 2
        Call opc05_2_2_Click 'opc1_01_Click
    Case 3
        Call opc4_2_Click  'opc1_03_Click
    Case 4
        Call mnurepger_Click
    Case 5
        Call opc1_01_Click
    Case 6
        Call opc05_3_1_Click
     
End Select
End Sub


Private Sub MDIForm_Load()
Dim rs As ADODB.Recordset

Set rs = VGCNx.Execute("select * from vt_configuraimpresora where " _
& " empresacodigo='" & VGParametros.empresacodigo & "' " _
& " and puntovtacodigo='" & VGParametros.puntovta & "' ")
If rs.RecordCount = 0 Then
    MsgBox "No existe impresora configurada." & Chr(13) & "Debe configurar impresora Matricial/Ticketera.", vbCritical, "ATENCION"
    FrmImpresora.Show
End If

EsFactura = False
VgModificar = False
Unload FrmIngreso
MostrarForm Me, "M"

End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    Set VGCNx = Nothing
    Set VGconfig = Nothing
    End
End Sub
Private Sub mnufv_Click()
FrmFv.Caption = "Formato de Venta"
FrmFv.Show
End Sub

Private Sub mnuhshshs_Click()
Dim Prod As String
Dim RsP As ADODB.Recordset

Set RsP = VGCNx.Execute("SELECT productocodigo,PRODUCTOPRECVTA FROM LISTAPRE1")
Do While Not RsP.EOF
    VGCNx.Execute "update maeart set aprecio=" & RsP!productoprecvta & " where acodigo='" & RsP!productocodigo & "'"
    RsP.MoveNext
Loop
End Sub

Private Sub mnuimpresora_Click()
FrmImpresora.Show
End Sub

Private Sub mnuingxcon_Click()
FrmFv.Caption = "Ingreso por contacto"
FrmFv.Show
End Sub

Private Sub mnunotas_Click()
    RptNotaCreRe.Show
End Sub

Private Sub mnunumeracion_Click()
     FrmRepCorr.Show
End Sub

Private Sub mnuprecio_Click()
    RptVariacionPrecio.Show
End Sub

Private Sub mnurankclie_Click()
    FrmRepRankCli.Show
End Sub

Private Sub mnurepger_Click()
FrmRepGer.Show
End Sub

Private Sub mnuresumen_Click()
  frmRepVtasMes.Show
End Sub

Private Sub mnutipcambio_Click()
    frmMantTipoCambio.Show
End Sub

Private Sub mnuvtaxfacdet_Click()
fmvtafacdet.Show
End Sub

Private Sub op4_10_01_Click()
FrmRepRankCli.Show
End Sub

Private Sub op4_10_05_Click()
FrmRepGer.Show
End Sub

Private Sub op4_10_06_Click()
frmRepVtasMes.Show
End Sub

Private Sub opc_priebas_Click()
Form1.Show 1
End Sub

Private Sub opc02_1_Click()
VgModificar = 0
Frmcorrecciondatosgen.Show 1
End Sub

Private Sub opc02_6_Click()
   FrmImportaData.Show
End Sub

Private Sub opc02_8_Click()
VgModificar = 1
Frmcorrecciondatosgen.Show 1
End Sub

Private Sub opc05_3_4_Click()
' FrmVentasxZonas.Show 1
End Sub

Private Sub opc05_7_Click()
FrmImpresora.Show
End Sub

Private Sub opc1_01_Click()
   FrmListaPrecios.Show
End Sub

Private Sub opc1_02_Click()
     FrmPedidoVentanilla.Show
End Sub

Private Sub opc1_03_Click()
   FrmAnulaFactura.Show
End Sub

Private Sub opc1_04_Click()
  FrmCopiaPedido.Show
End Sub

Private Sub opc1_05_Click()
    FrmImprimirguias.Show
End Sub

Private Sub Opc1_06_Click()
    FrmGeneracionpedidos.Show 1
End Sub

Private Sub opc1_07_Click()
VgModificar = False
FrmPedidoVentanilla.Show
End Sub

Private Sub opc1_08_Click()
VgModificar = True
   FrmPedidoVentanilla.Show
End Sub

Private Sub opc2_3_1_Click()
    FrmRepRankClientesNegocios.Show
End Sub

Private Sub opc2_2_Click()
  FrmCotizacionLibre.Show
End Sub

Private Sub opc2_3_Click()
    PrcEliminadocu.Show
End Sub

Private Sub opc2_4_Click()
    FrmTraslado.Show
End Sub

Private Sub opc2_5_Click()
  Frmtransfeclie.Show
End Sub

Private Sub opc3_1_Click()
  CstDocumentos.Show
End Sub

Private Sub opc3_2_Click()
  CstVentas.Show
End Sub

Private Sub opc3_3_Click()
  CstLimiteCredito.Show
End Sub

Private Sub opc4_1_1_Click()
    FrmRepVtasxArt.Show
End Sub

Private Sub opc4_1_2_Click()
    FrmRepVtasxFact.Show
End Sub

Private Sub opc4_1_3_Click()
    FrmRepGuiaFactBol.Show
End Sub

Private Sub opc4_1_9_Click()
  FrmClienteProdu.Show
End Sub


Private Sub opc3_4_Click()
FrmConPro.Show
End Sub

Private Sub opc3_5_Click()
FrmAnalisisVentas.Show
End Sub

Private Sub opc4_01_01_Click()
FrmRepVtasxArt.Show 1
End Sub

Private Sub opc4_01_04_Click()
 FrmComisionesVendedores.Show 1
End Sub

Private Sub opc4_01_08_Click()
FrmRepVtasxForma.Show
End Sub

Private Sub opc4_01_10_Click()
FrmRepDocAnula.Show
End Sub

Private Sub opc4_01_11_Click()
FrmRepCorr.Show
End Sub

Private Sub opc4_01_12_Click()
frmvtafacdet.Show
End Sub

Private Sub opc4_04_Click()
RptNotaCreRe.Show
End Sub

Private Sub opc4_06_Click()
FrmRepRegVtas.Show
End Sub

Private Sub opc4_08_Click()
FrmRepVtasxForma.Show
End Sub

Private Sub opc4_10_01_Click()
FrmRepRankCli.Show
End Sub

Private Sub opc4_10_02_Click()
FrmRepRankArt.Show
End Sub

Private Sub opc4_10_03_Click()
 frmRepRankvend.Show 1
End Sub

Private Sub opc4_10_05_Click()
FrmRepGer.Show
End Sub

Private Sub opc4_10_06_Click()
frmRepVtasMes.Show
End Sub

Private Sub opc4_2_Click()
 FrmRepRegVtas.Show
End Sub


Private Sub opc05_1_1_Click()
   FrmEmpresa.Show
End Sub

Private Sub opc05_1_2_Click()
  FrmAlmacen.Show
End Sub

Private Sub opc05_1_3_Click()
  frmBanco.Show
End Sub

Private Sub opc05_1_4_Click()
  FrmGrupoVenta.Show
End Sub

Private Sub opc05_1_5_Click()
 FrmFormaPago.Show
End Sub

Private Sub opc5_1_6_Click()
  FrmMoneda.Show
End Sub

Private Sub opc05_1_7_Click()
 FrmNegocio.Show
End Sub

Private Sub opc05_1_8_Click()
  FrmPuntoVenta.Show
End Sub

Private Sub opc05_1_9_Click()
  FrmUnidadMedida.Show
End Sub

Private Sub opc05_1_A_Click()
 FrmVendedor.Show 1
End Sub

Private Sub opc05_1_B_Click()
 FrmZona.Show 1
End Sub

Private Sub opc05_1_c_Click()
    FrmDocumento.Show 1
End Sub

Private Sub opc05_1_d_Click()
frmConceptosdePago.Show 1
End Sub

Private Sub opc05_2_2_Click()
 Frmcliente.Show
End Sub

Private Sub opc05_2_3_Click()
 FrmModoVenta.Show
End Sub

Private Sub opc05_2_4_Click()
 FrmParametroVenta.Show
End Sub

Private Sub opc05_2_5_Click()
 FrmLimiteCredito.Show
End Sub

Private Sub opc05_3_1_Click()
 FrmPtoVtaDoc.Show
End Sub

Private Sub opc05_3_2_Click()
 FrmSerieDocumento.Show
End Sub

Private Sub opc05_3_3_Click()
 FrmZonaVendedor.Show
End Sub

Private Sub opc05_5_Click()
vgtipo = facturacion
FrmCfgEmpresa.Show 1
End Sub

Private Sub opc05_6_Click()
vgtipo = facturacion
frmCfgUsuario.Show 1
End Sub

Private Sub opc6_Click()
   If MsgBox("Desea Salir del Sistema?", vbYesNo + vbQuestion, "AVISO") = vbYes Then
      Set VGCNx = Nothing
      Set Cn = Nothing
      Set cg = Nothing
      Set cnconta = Nothing
      End
   End If
End Sub

Private Sub Panel_PanelClick(ByVal Panel As MSComctlLib.Panel)
  If Panel.Index = 5 Then
     Load FrmIngreso
     FrmIngreso.Show 1
  End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

End Sub
