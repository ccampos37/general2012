VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.MDIForm MDIMain 
   BackColor       =   &H8000000C&
   Caption         =   "SISTEMA DE FACTURACION"
   ClientHeight    =   5400
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   8310
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport oCrystalReport 
      Left            =   360
      Top             =   4440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Menu mnuMaestros 
      Caption         =   "&Maestros"
      Begin VB.Menu mnuproducto 
         Caption         =   "Pro&ductos"
      End
      Begin VB.Menu mnumodoventa 
         Caption         =   "Modo &Venta"
      End
      Begin VB.Menu mnuparamvta 
         Caption         =   "Parametros Venta"
      End
      Begin VB.Menu mnuClientes 
         Caption         =   "&Clientes"
      End
      Begin VB.Menu mnulimcred 
         Caption         =   "Limite Crédito"
      End
   End
   Begin VB.Menu mnuBasicas 
      Caption         =   "&Basicas"
      Begin VB.Menu mnualmacen 
         Caption         =   "Almacen"
      End
      Begin VB.Menu mnubanco 
         Caption         =   "Banco"
      End
      Begin VB.Menu mnudocumento 
         Caption         =   "Documento"
      End
      Begin VB.Menu mnuEmpresa 
         Caption         =   "Empresa"
      End
      Begin VB.Menu mnuFormapago 
         Caption         =   "Forma Pago"
      End
      Begin VB.Menu mnugrupoventa 
         Caption         =   "Grupo Venta"
      End
      Begin VB.Menu mnumoneda 
         Caption         =   "Moneda"
      End
      Begin VB.Menu mnunegocio 
         Caption         =   "Negocio"
      End
      Begin VB.Menu mnuperfil 
         Caption         =   "Perfil"
      End
      Begin VB.Menu mnupuntoventa 
         Caption         =   "Punto Venta"
      End
      Begin VB.Menu mnutransaccion 
         Caption         =   "Transaccion"
      End
      Begin VB.Menu mnuunidad 
         Caption         =   "Unidad"
      End
      Begin VB.Menu mnuusuario 
         Caption         =   "Usuario"
      End
      Begin VB.Menu mnuvendedor 
         Caption         =   "Vendedor"
      End
      Begin VB.Menu mnuzona 
         Caption         =   "Zona"
      End
   End
   Begin VB.Menu mnureportes 
      Caption         =   "R&eportes"
      Begin VB.Menu mnumovvtas 
         Caption         =   "Movimientos de Venta"
         Begin VB.Menu mnuvtasxart 
            Caption         =   "Ventas x Articulo"
         End
         Begin VB.Menu mnuvtaxfact 
            Caption         =   "Ventas x Factura"
         End
         Begin VB.Menu mnuguiafactbol 
            Caption         =   "Guias,Facturas,Boletas"
         End
      End
      Begin VB.Menu mnurepcont 
         Caption         =   "Reportes Contables"
         Begin VB.Menu mnuregvtas 
            Caption         =   "Registro de Ventas"
         End
      End
      Begin VB.Menu mnuranking 
         Caption         =   "Ranking de Ventas"
         Begin VB.Menu mnurankart 
            Caption         =   "Ranking de Articulos"
         End
      End
   End
   Begin VB.Menu mnurelacion 
      Caption         =   "&Relacion"
      Begin VB.Menu mnupuntovtadoc 
         Caption         =   "Punto Venta - Documento"
      End
      Begin VB.Menu mnuseriedocumento 
         Caption         =   "Serie - Documento"
      End
      Begin VB.Menu mnuzonavendedor 
         Caption         =   "Zona - Vendedor"
      End
   End
   Begin VB.Menu mnuSalir 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "MDIMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub mnualmacen_Click()
    FrmAlmacen.Show
    FrmAlmacen.SetFocus
End Sub

Private Sub mnubanco_Click()
    frmBanco.Show
    frmBanco.SetFocus
End Sub

Private Sub mnucliente_Click()
  Frmcliente.Show
End Sub

Private Sub mnuClientes_Click()
    Frmcliente.Show
    Frmcliente.SetFocus
End Sub

Private Sub mnudocumento_Click()
    FrmDocumento.Show
    FrmDocumento.SetFocus
End Sub

Private Sub mnuEmpresa_Click()
    FrmEmpresa.Show
    FrmEmpresa.SetFocus
End Sub

Private Sub mnuFormapago_Click()
    FrmFormaPago.Show
    FrmFormaPago.SetFocus
End Sub

Private Sub mnugrupoventa_Click()
    FrmGrupoVenta.Show
    FrmGrupoVenta.SetFocus
End Sub

Private Sub mnuguiafactbol_Click()
    FrmRepGuiaFactBol.Show
    FrmRepGuiaFactBol.SetFocus
End Sub

Private Sub mnulimcred_Click()
    FrmLimiteCredito.Show
    FrmLimiteCredito.SetFocus
End Sub

Private Sub mnumodoventa_Click()
    FrmModoVenta.Show
    FrmModoVenta.SetFocus
End Sub

Private Sub mnumoneda_Click()
    FrmMoneda.Show
    FrmMoneda.SetFocus
End Sub

Private Sub mnunegocio_Click()
    FrmNegocio.Show
    FrmNegocio.SetFocus
End Sub

Private Sub mnupedido_Click()
    'FrmPedido.Show
End Sub

Private Sub mnuparamvta_Click()
    FrmParametroVenta.Show
    FrmParametroVenta.SetFocus
End Sub

Private Sub mnuperfil_Click()
    FrmPerfil.Show
    FrmPerfil.SetFocus
End Sub

Private Sub mnuproducto_Click()
    FrmProducto.Show
    FrmProducto.SetFocus
End Sub

Private Sub mnupuntoventa_Click()
    FrmPuntoVenta.Show
    FrmPuntoVenta.SetFocus
End Sub

Private Sub mnupuntovtadoc_Click()
    FrmPtoVtaDoc.Show
    FrmPtoVtaDoc.SetFocus
End Sub

Private Sub mnurankart_Click()
    frmRepRankArt.Show
    frmRepRankArt.SetFocus
End Sub

Private Sub mnuregvtas_Click()
    FrmRepRegVtas.Show
    FrmRepRegVtas.SetFocus
End Sub

Private Sub mnuSalir_Click()
End
End Sub

Private Sub mnuseriedocumento_Click()
    FrmSerieDocumento.Show
    FrmSerieDocumento.SetFocus
End Sub

Private Sub mnutransaccion_Click()
    FrmTransaccion.Show
    FrmTransaccion.SetFocus
End Sub

Private Sub mnuunidad_Click()
    FrmUnidadMedida.Show
    FrmUnidadMedida.SetFocus
End Sub

Private Sub mnuusuario_Click()
    FrmUsuario.Show
    FrmUsuario.SetFocus
End Sub

Private Sub mnuvendedor_Click()
    FrmVendedor.Show
    FrmVendedor.SetFocus
End Sub

Private Sub mnuvtasxart_Click()
    FrmRepVtasxArt.Show
    FrmRepVtasxArt.SetFocus
End Sub

Private Sub mnuvtaxfact_Click()
    FrmRepVtasxFact.Show
    FrmRepVtasxFact.SetFocus
End Sub

Private Sub mnuzona_Click()
    FrmZona.Show
    FrmZona.SetFocus
End Sub

Private Sub mnuzonavendedor_Click()
    FrmZonaVendedor.Show
    FrmZonaVendedor.SetFocus
End Sub
