VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form MDIPrincipal 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Sistema de Logistica y Almacenes"
   ClientHeight    =   8655
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   14745
   FillColor       =   &H00E0E0E0&
   ForeColor       =   &H8000000A&
   Icon            =   "Frmfox.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8655
   ScaleWidth      =   14745
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6930
      Top             =   6525
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frmfox.frx":15DBA
            Key             =   "Entrar"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frmfox.frx":28AC3
            Key             =   "Retornar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frmfox.frx":3B218
            Key             =   "Camara"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frmfox.frx":3BB45
            Key             =   "Tabla"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   8280
      Width           =   14745
      _ExtentX        =   26009
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6174
            MinWidth        =   6174
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Crystal.CrystalReport CryRptProc 
      Left            =   5670
      Top             =   6660
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   7590
      Top             =   6480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frmfox.frx":3C36D
            Key             =   "Facturar"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frmfox.frx":45905
            Key             =   "Modificar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frmfox.frx":54D5E
            Key             =   "Nuevo"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frmfox.frx":5F9BF
            Key             =   "Eliminar"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frmfox.frx":72CD3
            Key             =   "Facturado"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frmfox.frx":7D610
            Key             =   "Retornar"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frmfox.frx":8FD65
            Key             =   "Insertar"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frmfox.frx":99B82
            Key             =   "Sacar"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frmfox.frx":A39FC
            Key             =   "Grabar"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frmfox.frx":B2FBC
            Key             =   "Adicionar"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frmfox.frx":BD954
            Key             =   "Imprimir"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frmfox.frx":D3568
            Key             =   "Consultar"
         EndProperty
      EndProperty
   End
   Begin Crystal.CrystalReport oCrystalReport 
      Left            =   6255
      Top             =   6705
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Shape Linea 
      BackStyle       =   1  'Opaque
      Height          =   105
      Left            =   -45
      Top             =   0
      Width           =   20235
   End
   Begin VB.Menu Menu01 
      Caption         =   "&Mantenimiento"
      Begin VB.Menu Menu01_01 
         Caption         =   "&Artículos"
      End
      Begin VB.Menu menu01_02 
         Caption         =   "&Logística"
         Visible         =   0   'False
         Begin VB.Menu menu01_02_01 
            Caption         =   "Mant. Logistico"
         End
         Begin VB.Menu menu01_02_02 
            Caption         =   "Estado Aprobacion Requerimientos"
         End
         Begin VB.Menu menu01_02_03 
            Caption         =   "Estados Aprobacion Ordenes"
         End
      End
      Begin VB.Menu menu01_03 
         Caption         =   "&Proveedores"
         Visible         =   0   'False
      End
      Begin VB.Menu Menu01_04 
         Caption         =   "&Clientes"
         Visible         =   0   'False
      End
      Begin VB.Menu Men_mnu_alma 
         Caption         =   "Al&macenes"
      End
      Begin VB.Menu Men_ManTra 
         Caption         =   "Tra&nsacciones"
      End
      Begin VB.Menu Men_mnucasillero 
         Caption         =   "Ubicación de Articulos"
         Visible         =   0   'False
      End
      Begin VB.Menu Men_mnutransn 
         Caption         =   "&Transportista"
      End
      Begin VB.Menu Men_Kits 
         Caption         =   "&Mantenimiento de Kits"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_manten_lote_01 
         Caption         =   "Mantenimiento de Lotes"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu Men_CearCod 
         Caption         =   "&Crear Codigo de Tela"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu Men_CearCodHilo 
         Caption         =   "&Crear Codigo de Hilo"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu Men_CearCodAvio 
         Caption         =   "&Crear Codigo de Avios"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu Men_ManAyu 
         Caption         =   "Tablas de Ayudas"
         Begin VB.Menu mnu_docum_01 
            Caption         =   "Documentos"
         End
         Begin VB.Menu mnu_unidades_02 
            Caption         =   "Unidades de Medida"
         End
         Begin VB.Menu Men_ayuFam_03 
            Caption         =   "Familia de Artículos"
         End
         Begin VB.Menu mnu_defdoc_04 
            Caption         =   "Def. Documentos"
         End
         Begin VB.Menu mnu_auto_05 
            Caption         =   "Autorizado"
         End
         Begin VB.Menu mnu_tipArt_06 
            Caption         =   "Tipo de Articulo"
         End
         Begin VB.Menu mnu_giropro_07 
            Caption         =   "Giro de Proveedor"
            Visible         =   0   'False
         End
         Begin VB.Menu mnu_ccostos_08 
            Caption         =   "Centro de Costos"
            Visible         =   0   'False
         End
         Begin VB.Menu mnu_Distrito_09 
            Caption         =   "Distritos"
            Enabled         =   0   'False
            Visible         =   0   'False
         End
         Begin VB.Menu mnu_tipcam_10 
            Caption         =   "Tipo de Cambio"
            Enabled         =   0   'False
            Visible         =   0   'False
         End
         Begin VB.Menu mnu_clase_11 
            Caption         =   "Clase de Articulo"
            Enabled         =   0   'False
            Visible         =   0   'False
         End
         Begin VB.Menu mnu_color_12 
            Caption         =   "Color de Articulo"
            Enabled         =   0   'False
            Visible         =   0   'False
         End
         Begin VB.Menu mnu_ubica13 
            Caption         =   "Ubicaciones"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuCondPago_14 
            Caption         =   "Condición de Pago"
            Visible         =   0   'False
         End
         Begin VB.Menu mnu_Listatallas 
            Caption         =   "Lista de tallas"
            Enabled         =   0   'False
            Visible         =   0   'False
         End
         Begin VB.Menu Men_Tela 
            Caption         =   "Telas -->"
            Enabled         =   0   'False
            Visible         =   0   'False
            Begin VB.Menu Men_FamTela_15 
               Caption         =   "Familia de Tela"
            End
            Begin VB.Menu Men_TituTela_16 
               Caption         =   "Titulo de Tela"
            End
            Begin VB.Menu Men_MezclaTela_17 
               Caption         =   "Mezcla de Tela"
            End
            Begin VB.Menu Men_AnchoTela_18 
               Caption         =   "Ancho de Tela"
            End
            Begin VB.Menu Men_DensiTela_19 
               Caption         =   "Densidad de Tela"
            End
         End
         Begin VB.Menu Men_Avios 
            Caption         =   "Avios --->"
            Enabled         =   0   'False
            Visible         =   0   'False
            Begin VB.Menu Men_FamiliaAvios 
               Caption         =   "Familia de Avios"
            End
            Begin VB.Menu Men_OrigenAvios 
               Caption         =   "Origen de Avios"
               Enabled         =   0   'False
            End
            Begin VB.Menu Men_CalidadAvios 
               Caption         =   "Calidad de Avios"
            End
            Begin VB.Menu Men_CaractAvios 
               Caption         =   "Caracteristicas de Avi."
            End
            Begin VB.Menu Men_MedidaAvios 
               Caption         =   "Medida de Avios"
            End
            Begin VB.Menu Men_ColorAvios 
               Caption         =   "Color de Avios"
            End
         End
         Begin VB.Menu mnu_lineas 
            Caption         =   "Lineas"
         End
         Begin VB.Menu mnu_grupos 
            Caption         =   "Grupos"
         End
         Begin VB.Menu mnu_01_03 
            Caption         =   "Solicitantes"
         End
         Begin VB.Menu mnu_03_03 
            Caption         =   "Tipo de Maquinas"
         End
         Begin VB.Menu mnu_01_03_tipooc 
            Caption         =   "Tipo de orden de Compra"
         End
      End
      Begin VB.Menu Men_0106 
         Caption         =   "Entregas x Cliente"
      End
   End
   Begin VB.Menu mnulistado 
      Caption         =   "&Listados"
      Begin VB.Menu mnu_catarticulo 
         Caption         =   "Católogo de Artículo"
      End
      Begin VB.Menu mnu_catproveed 
         Caption         =   "Católogo de Proveedores"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu menu03 
      Caption         =   "R&equerimientos"
      Visible         =   0   'False
      Begin VB.Menu menu03_01 
         Caption         =   "Ing. Requerimientos PEDIDOS"
      End
      Begin VB.Menu menu03_02 
         Caption         =   "Requerimientos a LOGISTICA"
         Begin VB.Menu menu03_02_01 
            Caption         =   "Ingresos"
         End
         Begin VB.Menu menu03_02_02 
            Caption         =   "Primera Aprobacion"
         End
         Begin VB.Menu menu03_02_03 
            Caption         =   "Aprobacion Gerencia"
         End
         Begin VB.Menu menu03_02_04 
            Caption         =   "Reportes"
         End
         Begin VB.Menu menu03_02_05 
            Caption         =   "Seguimiento de Requerimientos"
         End
      End
   End
   Begin VB.Menu menu04 
      Caption         =   "M&ovimientos"
      Begin VB.Menu menu04_01 
         Caption         =   "Nota de Ingreso"
         Shortcut        =   ^I
      End
      Begin VB.Menu menu04_02 
         Caption         =   "Nota de Salidas "
         Shortcut        =   ^S
      End
      Begin VB.Menu menu04_03 
         Caption         =   "Guías de Remisión"
         Shortcut        =   ^G
      End
      Begin VB.Menu menu04_04 
         Caption         =   "Generacion  Liq. Ccompras"
         Visible         =   0   'False
      End
      Begin VB.Menu menu04_05 
         Caption         =   "Traslados"
      End
      Begin VB.Menu mnu04_06 
         Caption         =   "Transformacion"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu04_07 
         Caption         =   "Desdoble"
         Visible         =   0   'False
      End
      Begin VB.Menu menu04_08 
         Caption         =   "Emison de Orden de compra"
         Index           =   1
      End
      Begin VB.Menu menu04_09 
         Caption         =   "Ingreso por Orde. Compras/Requerimiento"
      End
      Begin VB.Menu menu04_10 
         Caption         =   "Armado de kits"
         Visible         =   0   'False
      End
      Begin VB.Menu menu04_11 
         Caption         =   "Desarmado de kits"
         Visible         =   0   'False
      End
      Begin VB.Menu menu04_12 
         Caption         =   "Emision de Liq. de Compras"
         Visible         =   0   'False
      End
      Begin VB.Menu menu04_13 
         Caption         =   "Ingresos"
      End
   End
   Begin VB.Menu menu07 
      Caption         =   "Informacion Global"
      Begin VB.Menu menu07_01 
         Caption         =   "Pedidos a Produccion"
         Visible         =   0   'False
      End
      Begin VB.Menu mnureque 
         Caption         =   "Requerimientos de Productos"
         Visible         =   0   'False
      End
      Begin VB.Menu menu07_04 
         Caption         =   "Saldos Consolidados x articulo"
      End
      Begin VB.Menu menu07_01_05 
         Caption         =   "Saldos Consolidados x Familia"
      End
   End
   Begin VB.Menu menu05 
      Caption         =   "&Consultas"
      Begin VB.Menu menu07_03 
         Caption         =   "Saldos Consolidados"
      End
      Begin VB.Menu mnu_stkArt1 
         Caption         =   "Stock de Artículos"
      End
      Begin VB.Menu mnu_conValArtPend 
         Caption         =   "Documentos"
      End
      Begin VB.Menu mnu_provart 
         Caption         =   "Proveedor por Artículo"
      End
      Begin VB.Menu mnu_docvalorizado 
         Caption         =   "Consulta de Doc. Valorizados"
      End
      Begin VB.Menu mnu_movart 
         Caption         =   "Resumen Movimiento Artículo Anual"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_movarti 
         Caption         =   "Movimiento por Articulo"
      End
      Begin VB.Menu menurep_04 
         Caption         =   "Requerimientos/ordenes de compra"
         Visible         =   0   'False
      End
      Begin VB.Menu menu05_08 
         Caption         =   "Saldos Totales"
         Visible         =   0   'False
      End
      Begin VB.Menu menu05_09 
         Caption         =   "kardex de articulo"
      End
   End
   Begin VB.Menu men 
      Caption         =   "&Reportes"
      Begin VB.Menu Men_RepAlm 
         Caption         =   "Almacén"
         Begin VB.Menu Men_AlmStock_01 
            Caption         =   "Stock de Artículos"
         End
         Begin VB.Menu Men_AlmStock_10 
            Caption         =   "Stock de Artículos x Fecha"
         End
         Begin VB.Menu Men_AlmStock_02 
            Caption         =   "Stock por Lote/Serie"
         End
         Begin VB.Menu Men_AlmKar_02 
            Caption         =   "Kardex de Artículos"
         End
         Begin VB.Menu mnu_repdocAlm 
            Caption         =   "Documentos del Almacen"
            Enabled         =   0   'False
         End
         Begin VB.Menu Men_InvMovKar_03 
            Caption         =   "Kardex de Lote"
         End
         Begin VB.Menu mnu_artven_06 
            Caption         =   "Artículos Vencidos"
         End
         Begin VB.Menu mnu_artven_08 
            Caption         =   "Informe de Stock Inicial"
         End
         Begin VB.Menu mnu_maeart_10 
            Caption         =   "Maestro de Articulo"
            Visible         =   0   'False
         End
      End
      Begin VB.Menu Men_RepVal 
         Caption         =   "Inventario Valorado x Almacen"
         Begin VB.Menu Men_InvKarVal_01 
            Caption         =   "Kardex Valorizado Resumen"
         End
         Begin VB.Menu Men_InvKarVal_02 
            Caption         =   "Kardex Valorizado Detallado"
         End
         Begin VB.Menu mnu_kdxcencos_02 
            Caption         =   "kardex Valorizado por Centro Costo"
         End
         Begin VB.Menu mnu_valxdoc_03 
            Caption         =   "Kardex Valorizado por Documentos"
         End
         Begin VB.Menu menu_val_05 
            Caption         =   "Kardex Valorizado por Transaccion"
         End
      End
      Begin VB.Menu mnu_repdoc 
         Caption         =   "Documentos"
         Begin VB.Menu nmu_doc01 
            Caption         =   "Detallado"
         End
         Begin VB.Menu mnu_doc02 
            Caption         =   "Resumido"
         End
         Begin VB.Menu mnu_Trasla 
            Caption         =   "Informe de Traslados"
         End
         Begin VB.Menu mnu_doc04 
            Caption         =   "Ordenes de Fabricacion"
            Visible         =   0   'False
         End
         Begin VB.Menu mnu_repdoc_05 
            Caption         =   "Comparativos Almacenes - ventas"
         End
      End
      Begin VB.Menu menu_0307 
         Caption         =   "Otros"
         Begin VB.Menu mnu_consxcc_04 
            Caption         =   "Consumo x Centro Costos"
         End
         Begin VB.Menu mnu_articulocxcosto_04 
            Caption         =   "Consumo Articulos en Centro de Costos"
         End
         Begin VB.Menu mnu_artven_07 
            Caption         =   "Utilidades Venta/Costo"
            Visible         =   0   'False
         End
         Begin VB.Menu menu_0307_05 
            Caption         =   "Ingresos x Proveedor"
         End
         Begin VB.Menu menurep_05_01 
            Caption         =   "Ordenes de Compra"
            Visible         =   0   'False
         End
         Begin VB.Menu mnu_articulos 
            Caption         =   "Articulos Sin Movimientos"
         End
      End
      Begin VB.Menu men_ValxEstab 
         Caption         =   "Inventario Valorizado x establecimiento"
         Visible         =   0   'False
         Begin VB.Menu men_ValxEstab_02 
            Caption         =   "Kardex Valorizado detallado"
         End
      End
      Begin VB.Menu men_ValxEmp 
         Caption         =   "Inventario Valorado x  Empresa"
         Begin VB.Menu men_ValxEmp_01 
            Caption         =   "Kardex Valorizado Resumen"
         End
         Begin VB.Menu men_ValxEmp_02 
            Caption         =   "kardex Valorizado Detallado"
         End
         Begin VB.Menu men_ValxEmp_03 
            Caption         =   "Kardex Valorizado Transaccion"
         End
      End
   End
   Begin VB.Menu Pro 
      Caption         =   "&Procesos"
      Begin VB.Menu Men_ProVal 
         Caption         =   "Valorización"
         Begin VB.Menu Men_TraVal 
            Caption         =   "Valorización Art. Pendientes"
         End
         Begin VB.Menu menu08_02_02 
            Caption         =   "Valorizacion ( inc. gastos)"
            Visible         =   0   'False
         End
         Begin VB.Menu Men_TraCor 
            Caption         =   "Correción Art Valorizados Masivo"
            Visible         =   0   'False
         End
         Begin VB.Menu mnu_contrans_06_01 
            Caption         =   "Valorización Mensual"
         End
         Begin VB.Menu menu08_01_05 
            Caption         =   "Valorizacion x empresa"
            Visible         =   0   'False
         End
         Begin VB.Menu menu08_01_06 
            Caption         =   "Valorizacion x establecimiento"
            Visible         =   0   'False
         End
         Begin VB.Menu mnu_genera 
            Caption         =   "Genera consumos ( temporal )"
            Visible         =   0   'False
         End
      End
      Begin VB.Menu pro_Esp 
         Caption         =   "Especiales"
         Visible         =   0   'False
         Begin VB.Menu pro_Esp_InvFis 
            Caption         =   "Inventario Físico"
            Begin VB.Menu pro_EspInvFis_01_01 
               Caption         =   "Registro de Existencias"
            End
            Begin VB.Menu pro_EspInvFis_01_02 
               Caption         =   "Informes de Inventario Fisico"
            End
         End
      End
      Begin VB.Menu Men_GuiRem 
         Caption         =   "Guia de Remisión"
         Begin VB.Menu Men_GuiEli_01 
            Caption         =   "Anular"
         End
         Begin VB.Menu Men_GuiDev_02 
            Caption         =   "Devolver"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu Men_GuiDoc 
         Caption         =   "Documentos"
         Visible         =   0   'False
         Begin VB.Menu Men_CocMod_01 
            Caption         =   "Modificar"
         End
         Begin VB.Menu Men_EliDoc_02 
            Caption         =   "Eliminar"
         End
      End
      Begin VB.Menu mnlinea1_02 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_ajuste 
         Caption         =   "Ajuste de Valorización"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_recstk_03 
         Caption         =   "Recalculo Stock"
         Begin VB.Menu mnu_contrans_08 
            Caption         =   "Recalculo Saldo Fisico"
         End
         Begin VB.Menu mnu_recstk_03_01 
            Caption         =   "Recalculo de Stock por Articulos"
            Visible         =   0   'False
         End
         Begin VB.Menu mnu_recstk_03_02 
            Caption         =   "Recalculo de Stock por Series/Lotes"
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mnu_contrans_09 
         Caption         =   "Anulacion de Documentos"
         Begin VB.Menu mnu_contrans_09_01 
            Caption         =   "Documentos"
         End
         Begin VB.Menu mnu_contrans_09_02 
            Caption         =   "Liquidacion de compras"
            Visible         =   0   'False
         End
         Begin VB.Menu mnu_contrans_09_03 
            Caption         =   "Transferencias"
         End
      End
      Begin VB.Menu menu_07_10 
         Caption         =   "Modifica traslados"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnu_auditor01 
      Caption         =   "Auditoria"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnu_contrans_06 
         Caption         =   "Revalorización"
         Begin VB.Menu mnu_contrans_06_03 
            Caption         =   "Cierre Mensual"
         End
      End
      Begin VB.Menu mn_linea_03 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_sincroTC 
         Caption         =   "Actualiza Tipo Cambio"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
   Begin VB.Menu menu09 
      Caption         =   "Confi&guración"
      Begin VB.Menu menu09_01 
         Caption         =   "Empresas"
         Visible         =   0   'False
      End
      Begin VB.Menu menu09_02 
         Caption         =   "Usuarios"
      End
      Begin VB.Menu Men_SisAdminis_03 
         Caption         =   "Admnistradores"
         Visible         =   0   'False
      End
      Begin VB.Menu menu09_03 
         Caption         =   "Parámetros"
         Visible         =   0   'False
      End
      Begin VB.Menu sisra_03 
         Caption         =   "-"
      End
      Begin VB.Menu Men_SisCam 
         Caption         =   "Cambiar Empresa"
         Shortcut        =   {F12}
         Visible         =   0   'False
      End
      Begin VB.Menu Men_SisAl 
         Caption         =   "Cambiar Almacén"
         Enabled         =   0   'False
         Shortcut        =   {F11}
         Visible         =   0   'False
      End
      Begin VB.Menu mnumigra 
         Caption         =   "Pruebas"
      End
   End
   Begin VB.Menu menu10 
      Caption         =   "Salir"
   End
End
Attribute VB_Name = "MDIPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adoreg As ADODB.Recordset
Dim rs As ADODB.Recordset

Private Sub Cmd1_Click()
FrmArArticulo.Show 1
End Sub

Private Sub Cmd2_Click()
  FrmAlmacen.Show 1
End Sub

Private Sub Cmd3_Click()
   VGRegEnt = 1
   FrmRegistro.Show
End Sub

Private Sub Cmd4_Click()
  VGGuiaSal = True
  VGRegEnt = 2
  FrmGuiaSal.Show
End Sub

Private Sub Cmd5_Click()
    FrmArmadoKits.Show
End Sub

Private Sub Cmd6_Click()
FrmFpp.Show
End Sub

Private Sub Command2_Click()

End Sub

Private Sub Cmd7_Click()
FrmSaldosConsolidados.Show
End Sub

Private Sub Cmd8_Click()
  FormConStk.Show 1
End Sub

Private Sub Cmd9_Click()
If MsgBox("Esta seguro que desea salir?", vbYesNo + vbInformation, "Sistemas") = vbYes Then End

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim i As Integer
For i = Forms.count - 1 To 0 Step -1
  Unload Forms(i)
Next
End Sub

Private Sub Men_0106_Click()
FrmEntregaxCliente.Show 1
End Sub

Private Sub Men_AlmKar_02_Click()
   FrmKardex.Show 1
End Sub
Private Sub Men_AlmStock_01_Click()
    FrmStockAlmacen.Show 1
   End Sub
Private Sub Men_AlmStock_02_Click()
frmStockLoteSerie.Show 1
End Sub
Private Sub Men_AlmStock_10_Click()
   FrmStockFecha.Show 1

End Sub
Private Sub Men_AnchoTela_18_Click()
FrmAnchoTela.Show 1
End Sub
Private Sub Men_ayuFam_03_Click()
   FrmMntFamilia.Show 1                           ' Form5.show 1
End Sub
Private Sub Men_CalidadAvios_Click()
FrmCalidadAvio.Show 1
End Sub
Private Sub Men_CaractAvios_Click()
FrmCaracteAvio.Show 1
End Sub
Private Sub Men_CearCod_Click()
FrmCreaCodigo.Show 1
End Sub
Private Sub Men_CearCodAvio_Click()
FrmCrearCodAvios.Show 1
End Sub
Private Sub Men_CearCodHilo_Click()
FrmCrearCodigoHilo.Show 1
End Sub
Private Sub Men_CocMod_01_Click()
  FrmModificar.Show 1
End Sub
Private Sub Men_ColorAvios_Click()
FrmColorAvio.Show 1
End Sub
Private Sub Men_DensiTela_19_Click()
FrmDesidadTela.Show 1
End Sub
Private Sub Men_EliDoc_02_Click()
VGElimina = True
'FormEliminaDoc.Show 1
'FormEliminaDoc.Show 1
End Sub
Private Sub pro_Esp_InvFis_01_Click()
   frmRegistroInventarioFisico.Show 1
   'frmtoma.show 1
End Sub
Private Sub pro_Esp_InvFis_02_Click()
'frmtoma.show 1
frmInformeInventarioFisico.Show 1
End Sub

Private Sub Form_Load()
Dim sFileName As String
Dim sBD As String
Dim sBDt As String
Dim n As String
Dim RSQL As String
Dim IASA As String
On Error GoTo Err
Set VGdllApi = New dll_apisgen.dll_apis
   
'Verificar_Sistema
VGCodMon = "01"
VGtransp = True
VGSALIR = False
VGcomputer = UCase(ComputerName)
VGsql = VGdllApi.LeerIni(App.Path & "\Marfice.ini", "conexion", "SQL", "")
VGsql = IIf(VGsql = "", 0, VGsql)
   
GPunto = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "CONEXION", "PUNTOVTA", "?")
GPunto = IIf(GPunto = "?", "01", GPunto)
g_ptoventa = GPunto

VGformatofecha = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "CONEXION", "FECHASQL", "?")
VGformatofecha = IIf(VGformatofecha = "?", "MDY", VGformatofecha)
       
'Conexion de General
VGParamSistem.BDEmpresaGEN = Trim(VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "BDGENERAL", "BDDATOS", "?"))
VGParamSistem.ServidorGEN = Trim(VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "BDGENERAL", "SERVIDOR", "?"))
VGParamSistem.UsuarioGEN = Trim(VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "BDGENERAL", "USUARIO", "?"))
' VGParamSistem.PwdGEN = DECODIFICA(Trim(VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "BDGENERAL", "PASSW", "?")), NUMMAGICO)
VGParamSistem.PwdGEN = Trim(VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "BDGENERAL", "PASSW", "?"))
        
'Conexion de inventarios
VGParamSistem.BDEmpresa = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "conexion", "BDDATOS", "?")
VGParamSistem.Servidor = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "conexion", "SERVIDOR", "?")
VGParamSistem.Usuario = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "conexion", "USUARIO", "?")
' VGParamSistem.PWD = DECODIFICA(Trim(VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "conexion", "PASSW", "?")), NUMMAGICO)
VGParamSistem.Pwd = Trim(VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "conexion", "PASSW", "?"))

VGOrden = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "conexion", "ORDEN", "?")
   
   ' reportes
VGParamSistem.RutaReport = Trim(VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "REPORTES", "INVENTARIOS", "?"))
VGParamSistem.carpetareportes = Trim(VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "conexion", "CARPETAREPORTES", "?"))
   
'Conexion de Contabilidad
VGParamSistem.BDEmpresaCT = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "CONTABILIDAD", "BDDATOS", "?")
If VGParamSistem.BDEmpresaCT = "" Then
   VGParamSistem.BDEmpresaCT = VGParamSistem.BDEmpresa
   VGParamSistem.ServidorCT = VGParamSistem.Servidor
   VGParamSistem.UsuarioCT = VGParamSistem.Usuario
   VGParamSistem.PwdCT = VGParamSistem.Pwd
Else
   VGParamSistem.BDEmpresaCT = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "CONTABILIDAD", "BDDATOS", "?")
   VGParamSistem.ServidorCT = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "CONTABILIDAD", "SERVIDOR", "?")
   VGParamSistem.UsuarioCT = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "CONTABILIDAD", "USUARIO", "?")
 '   VGParamSistem.PwdCT = DECODIFICA(Trim(VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "CONTABILIDAD", "PASSW", "?")), NUMMAGICO)
   VGParamSistem.PwdCT = Trim(VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "CONTABILIDAD", "PASSW", "?"))

End If

If VGParamSistem.RutaReport = "" Or VGParamSistem.RutaReport = "?" Then
   VGParamSistem.RutaReport = App.Path
   VGParamSistem.carpetareportes = "Reportes"
End If
       
'Establecer Cadena de Conexión de Reportes
VGCadenaReport2 = "DSN=jckconsultores;DSQ=" & VGParamSistem.BDEmpresaGEN & ";UID=" & VGParamSistem.UsuarioGEN & ";PWD=" & VGParamSistem.PwdGEN & ""
          
mensaje1 = "Prueba - Inventarios"
sFileName = "marfice.ini"
VGDIRE = sGetIni("Marfice.ini", "CONFIG", "DIRE", "?")

frmlogin.Show 1
MDIPrincipal.Caption = "Sistema de Inventario Empresa : " & VGParametros.NomEmpresa & "   Base de datos --> " & VGParamSistem.BDEmpresa

If VGSALIR Then
   If VGCNx.State = 1 Then VGCNx.Close
   If VGcnxCT.State = 1 Then VGcnxCT.Close
      MDIPrincipal.Visible = False
      Form_Unload (0)
      Exit Sub
Else
      Call ParametrosdeAlmacenes
End If

VGAutomatico = False
Linea.Width = Me.Width

Exit Sub

Err:
    MsgBox Err.Description, vbExclamation, "Aviso"
    Exit Sub
    Resume
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
End Sub

Private Sub Men_FamiliaAvios_Click()
FrmFamiliaAvios.Show 1
End Sub
Private Sub Men_FamTela_15_Click()
FrmFamTela.Show 1
End Sub
Private Sub Men_GuiDev_02_Click()
VGGuiaSal = False
FrmGuiaSal.Show 1
End Sub
Private Sub Men_GuiEli_01_Click()
VGElimina = False
'FormEliminaDoc.Show 1

End Sub
Private Sub InvArtVal_Click()
FormArtVal.Show 1
End Sub
Private Sub Men_InvKarVal_01_Click()
formkardexValResumen.Show 1
End Sub
Private Sub InvMovKar_Click()
FormKardexMov.Show 1
End Sub
Private Sub Men_InvKarVal_02_Click()
   FrmKardexValTXDetallado.Show 1
End Sub
Private Sub Men_InvMovKar_03_Click()
frmKardexLote.Show 1
End Sub
Private Sub Men_Kits_Click()
'FrmRegKit.show 1
FrmRegPlantilladeKits.Show 1
End Sub
Private Sub Men_ManArt_Click()
   FrmArArticulo.Show 1
End Sub
Private Sub Men_MantClie_Click()
'  VGAyuClie = False
  FrmArClien.Show 1
End Sub
Private Sub Men_MantPro_Click()
  Frmcliente.Show 1
End Sub
Private Sub Men_ManTra_Click()
  FrmTransaccion.Show 1
End Sub
Private Sub Men_MedidaAvios_Click()
FrmMedidaAvio.Show 1
End Sub
Private Sub Men_MezclaTela_17_Click()
FrmMezclaTela.Show 1
End Sub
Private Sub Men_mnGui_Click()
  VGGuiaSal = True
  VGRegEnt = 2
  FrmGuiaSal.Show 1
End Sub
Private Sub Men_mnu_alma_Click()
  FrmAlmacen.Show 1
End Sub
Private Sub Men_TituTela_16_Click()
FrmTituloTela.Show 1
End Sub

Private Sub men_ValxEmp_01_Click()
FrmKardexValEmpresa.Show 1
End Sub

Private Sub men_ValxEmp_02_Click()
FrmKarValdetxEmpresa.Show 1
End Sub

Private Sub men_ValxEmp_03_Click()
frmKarValTransxEmpresa.Show 1
End Sub

Private Sub men_ValxEstab_02_Click()
FrmKarValxEst.Show 1
End Sub

Private Sub menu_0307_05_Click()
FrmRepArtxProveedor.Show 1
End Sub
Private Sub menu_07_10_Click()
FrmModificaTraslados.Show
End Sub


Private Sub menu_val_05_Click()
FrmKardexValTransaccion.Show 1
End Sub

Private Sub Menu01_01_Click()
FrmArArticulo.Show 1
End Sub

Private Sub menu01_02_01_Click()
Frmlogistica.Show 1
End Sub

Private Sub menu01_02_02_Click()
FrmEstadoRequerimientos.Show 1
End Sub

Private Sub menu01_02_03_Click()
FrmEstadoOrdenes.Show 1
End Sub
Private Sub Menu01_04_Click()
Frmcliente.Show 1
End Sub

Private Sub menu03_02_01_Click()
 frmRequerimientosOrdenes.Show 1
End Sub

Private Sub menu03_02_02_Click()
VGtipoAprobacion = 0
VGOrdenes = 1
frmAprobacionRequerimientos.Show 1
End Sub

Private Sub menu03_02_03_Click()
VGtipoAprobacion = 1
VGOrdenes = 1
frmAprobacionRequerimientos.Show 1
End Sub
Private Sub menu03_01_Click()
frmRequerimientosPedidos.Show
End Sub

Private Sub menu03_04_Click()
 FrmRequerimientosReportes.Show 1
End Sub

Private Sub menu03_05_Click()
 FrmRequerimientoSeguimiento.Show 1
End Sub



Private Sub menu04_01_Click()
   VGRegEnt = 1
   FrmRegistro.Show
End Sub
Private Sub menu04_02_Click()
   VGRegEnt = 0
   FrmRegistro.Show
End Sub
Private Sub menu04_03_Click()
  VGGuiaSal = True
  VGRegEnt = 2
  FrmGuiaSal.Show
End Sub
Private Sub menu04_04_Click()
FrmGeneraLiqCompras.Show
End Sub
Private Sub menu04_05_Click()
FrmTraslado.Show
End Sub
Private Sub menu04_06_Click()
    frmOrdenes.Show 1
End Sub

Private Sub menu04_08_Click(Index As Integer)
If VGParametros.PermiteRequerimientos Then
    FrmOrdenes_Requerimientos.Show 1
 Else
    frmOrdenes.Show 1
End If
End Sub
Private Sub menu04_09_Click()
    frmingresoOC.Show
End Sub
Private Sub menu04_10_Click()
    FrmArmadoKits.Show
End Sub
Private Sub menu04_11_Click()
    FrmDesKits.Show
End Sub
Private Sub menu04_12_Click()
    FrmliqCompra.Show
End Sub

Private Sub menu04_13_Click()
   VGRegEnt = 1
FrmMntMovimientos.Show
End Sub

Private Sub menu05_08_Click()
FrmSaldostotales.Show 1
End Sub

Private Sub menu07_05_Click()
FrmSaldostotales.Show 1
End Sub

Private Sub menu05_09_Click()
FrmConsultakardex.Show
End Sub

Private Sub menu07_01_05_Click()
FrmSaldostotales.Show 1
End Sub



Private Sub menu08_01_05_Click()
FrmValorizacionxEmpresa.Show 1
End Sub

Private Sub menu08_01_06_Click()
FrmValorizacionxEstablecimiento.Show 1
End Sub

Private Sub menu08_02_02_Click()
   FrmValorizacionArticulos.Show 1
End Sub

Private Sub menu09_02_Click()
frmCfgUsuario.Show 1
End Sub

Private Sub menu10_Click()
Unload Me
End Sub

Private Sub menu07_01_Click()
FrmFpp.Show 1
End Sub

Private Sub menurep_04_Click()
FrmreporteOrdenesdecompra.Show 1
End Sub

Private Sub menurep_05_01_Click()
FrmreporteOrdenesdecompra.Show
End Sub

Private Sub mnu_01_03_Click()
    FrmMntSolicitantes.Show 1
End Sub

Private Sub mnu_01_03_tipooc_Click()
FrmMntTipodeOrden.Show 1
End Sub

Private Sub mnu_03_03_Click()
    FrmMntMaquinas.Show 1
End Sub
Private Sub mnu_articulocxcosto_04_Click()
    frmArticuloXCenCos.Show 1
End Sub

Private Sub mnu_articulos_Click()
frmArtSinMov.Show 1
End Sub

Private Sub mnu_artven_06_Click()
    FrmArtVen.Show 1
End Sub
Private Sub mnu_artxven_Click()
   FrmArtVen.Show 1
End Sub
Private Sub mnu_artven_07_Click()
    frmUtilidadVentaCosto.Show 1
End Sub
Private Sub mnu_artven_08_Click()
   FrmStockInicial.Show 1
End Sub
Private Sub mnu_Asiento_02_Click()
  FrmAsiento2001.Show 1
End Sub
Private Sub mnu_auto_05_Click()
  frmAutorizado.Show 1
End Sub
Private Sub Men_mnucasillero_Click()
  FormCasillero.Show 1
End Sub
Private Sub mnu_catarticulo_Click()
  FrmLisArticulos.Show 1
End Sub
Private Sub mnu_catproveed_Click()
  VGRclie = False
  FrmRepClie.Show 1
End Sub
Private Sub mnu_ccostos_08_Click()
  FrmCCostoTrans.Show 1
End Sub
Private Sub mnu_clase_11_Click()
  FrmArClase.Show 1
End Sub
Private Sub mnu_color_12_Click()
   FrmArColor.Show 1
End Sub
Private Sub mnu_consxcc_04_Click()
   VGcc = 1
   frmCenCos.Show 1
End Sub
Private Sub mnu_contrans_05_Click()
 'FrmConfdeTrans.show 1
 FrmConfdeTrans2.Show 1
End Sub
Private Sub mnu_contrans_06_01_Click()
  FrmKardexRevalorizaMes.Show 1
End Sub

Private Sub mnu_contrans_08_Click()
  FrmPrcSaldos.Show
End Sub

Private Sub mnu_contrans_09_01_Click()
VGtransf = 0
 frmAnulaDocumento.Show
End Sub

Private Sub mnu_contrans_09_02_Click()
FrmAnulaLiquidacionCompra.Show 1
End Sub

Private Sub mnu_contrans_09_03_Click()
VGtransf = 1
 frmAnulaDocumento.Show
End Sub

Private Sub mnu_conValArtPend_Click()
 FrmConsultaNotas.Show
End Sub
Private Sub mnu_defdoc_04_Click()
 FrmCfgDocumento.Show
End Sub
Private Sub mnu_Distrito_09_Click()
FrmArTabAyu.G_nOpc = 7
FrmArTabAyu.G_cTabla = "13"
FrmArTabAyu.Show 1
End Sub
Private Sub mnu_doc02_Click()
FrmRepDocuResumen.Show 1
End Sub

Private Sub mnu_doc04_Click()
FrmReporteOrdFabricacion.Show 1
End Sub

Private Sub mnu_docum_01_Click()
FrmArDocumento.Show 1
End Sub
Private Sub mnu_docvalorizado_Click()
FormConDocVal.Show 1
End Sub

Private Sub mnu_genera_Click()
xx_generaconsumos.Show
End Sub

Private Sub mnu_giropro_07_Click()
  FrmArTabAyu.G_nOpc = 10
  FrmArTabAyu.G_cTabla = "62"
  FrmArTabAyu.Show 1
End Sub
Private Sub mnu_grupos_Click()
   frmgrupo.Show 1
End Sub
Private Sub mnu_kdxcencos_02_Click()
  FrmKarValC.Show 1
End Sub
Private Sub mnu_Lineas_Click()
  Frmlineas.Show
End Sub
Private Sub mnu_Listatallas_Click()
    FrmListaTalla.Show 1
End Sub
Private Sub mnu_maeart_10_Click()
FrmRepMae.Show 1
End Sub
Private Sub mnu_manten_lote_01_Click()
   frmReglotes.Show 1
End Sub
Private Sub mnu_movart_Click()
  FormMovArt.Show 1
End Sub
Private Sub mnu_movarti_Click()
  cstkardexmovi.Show 1
End Sub
Private Sub mnu_provart_Click()
    frmPrueba.Show 1
End Sub

Private Sub mnu_recstk_03_02_Click()
frmCalcularLote_Serie2.Show 1
End Sub

Private Sub mnu_repdoc_05_Click()
FrmComparativos.Show
End Sub

Private Sub mnu_repdocAlm_Click()
   FrmRepMovFec.Show 1
End Sub
Private Sub mnu_RestSalAnt_04_Click()
 frmRestSalAnt.Show 1
End Sub

Private Sub mnu_sincroTC_Click()
 FrmSincronizaTC.Show 1
End Sub
Private Sub mnu_stkArt1_Click()
 FormConStk.Show 1
End Sub
Private Sub mnu_tipArt_06_Click()
  FrmArTipoArticulo.Show 1
End Sub
Private Sub mnu_tipcam_Click()
FrmArTipoCambio.Show 1
End Sub
Private Sub Men_mnutransn_Click()
 FrmTranspor.Show 1
End Sub
Private Sub mnu_tipcam_10_Click()
FrmArTipoCambio.Show 1
End Sub
Private Sub mnu_tool_001_Click()
'    UpdateDatabases.show 1
End Sub
Private Sub mnu_tool_003_Click()
'    FrmImpoExpo.show 1
End Sub
Private Sub mnu_tool_004_Click()
'    FrmImpo.show 1
End Sub
Private Sub mnu_trasla_Click()
  FrmRepTraslados.Show 1
End Sub
Private Sub mnu_Traslado_Click()
FrmTraslado.Show 1
End Sub
Private Sub mnu_ubica13_Click()
FrmMantenUbica.Show 1
End Sub
Private Sub mnu_unidades_02_Click()
 ' FrmArUniMed.show 1
 FrmMntUnidMedida.Show
 End Sub

Private Sub mnu_valxdoc_03_Click()
 RepKardexValTXDocumento.Show 1
End Sub
Private Sub Men_SisAdminis_03_Click()
FrmCfgAdminist.Show 1
End Sub

Private Sub Men_SisCam_Click()
FrmCfgCambioEmp.Show 1
End Sub
Private Sub Men_SisUsu_02_Click()
'  frmUsuario.show 1
  frmCfgUsuario.Show 1
End Sub
Private Sub Men_TraCor_Click()
  FrmCorrigeart.Show
End Sub
Private Sub Men_TraRegEnt_Click()
   VGRegEnt = 1
   FrmRegistro.Show 1
End Sub
Private Sub Men_TraRegSal_Click()
  VGRegEnt = 0
  FrmRegistro.Show 1
End Sub
Private Sub Men_TraVal_Click()
 FrmValArtPed.Show 1
End Sub

Private Sub mnu04_06_Click()
    FrmDistribucion.Show
End Sub
Private Sub mnu04_07_Click()
    FrmDistribucion_1.Show
End Sub

Private Sub mnureque_Click()
On Error GoTo errores

                      
Screen.MousePointer = 11
                                   
With oCrystalReport
        .Reset
        .ReportFileName = VGParamSistem.RutaReport & "al_requerimientos.rpt"

       If VGsql = 1 Then
           .Connect = "Provider=SQLOLEDB;Password=" & VGParamSistem.PwdGEN & ";Persist Security Info=True;User ID=" & VGParamSistem.UsuarioGEN & ";Initial Catalog=" & VGParamSistem.BDEmpresaGEN & ";SERVER=" & VGParamSistem.ServidorGEN
          Else
           .Connect = VGCadenaReport2
        End If

        .DiscardSavedData = True
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .WindowShowPrintSetupBtn = True
        .WindowShowExportBtn = True
        .WindowShowZoomCtl = True
        .WindowShowNavigationCtls = True
        .WindowShowPrintBtn = True
        .WindowTitle = "Formato Pedido Produccion"
        .StoredProcParam(0) = VGParamSistem.BDEmpresa
        '.StoredProcParam(1) = DtDesde
        '.StoredProcParam(2) = DtHasta
'        If Len(Trim(Ctr_almacen.xclave)) <> 0 Then
'            .StoredProcParam(1) = Trim(Ctr_almacen.xclave)
'        Else
'            .StoredProcParam(1) = "%"
'        End If
        
        .Action = 1
        
  End With
  
Screen.MousePointer = 1

Exit Sub
errores:
  Screen.MousePointer = 1
  MsgBox "Error inesperado: " & Err.Number & "  " & Err.Description, vbExclamation
  Err = 0
  Exit Sub
End Sub

Private Sub menu07_04_Click()
   FrmSaldosConsolidados.Show 1
End Sub

Private Sub nmu_doc01_Click()
 FrmDocuDeta.Show 1
End Sub


