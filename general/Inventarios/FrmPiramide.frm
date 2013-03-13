VERSION 5.00
Begin VB.Form FrmPrincipal 
   BackColor       =   &H8000000C&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sistema de Inventario"
   ClientHeight    =   6828
   ClientLeft      =   972
   ClientTop       =   1860
   ClientWidth     =   11568
   Icon            =   "FrmPiramide.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6828
   ScaleWidth      =   11568
   WindowState     =   2  'Maximized
   Begin VB.Menu mant 
      Caption         =   "&Mantenimeinto"
      Begin VB.Menu Men_ManArt 
         Caption         =   "&Artículos"
      End
      Begin VB.Menu Men_mnulogistica 
         Caption         =   "&Logística"
      End
      Begin VB.Menu Men_MantPro 
         Caption         =   "&Proveedores"
      End
      Begin VB.Menu Men_MantClie 
         Caption         =   "&Clientes"
         Enabled         =   0   'False
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
      End
      Begin VB.Menu Men_CearCodAvio 
         Caption         =   "&Crear Codigo de Avios"
         Enabled         =   0   'False
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
         End
         Begin VB.Menu mnu_ccostos_08 
            Caption         =   "Centro de Costos"
         End
         Begin VB.Menu mnu_Distrito_09 
            Caption         =   "Distritos"
         End
         Begin VB.Menu mnu_clase_11 
            Caption         =   "Clase de Articulo"
         End
         Begin VB.Menu mnu_ubica13 
            Caption         =   "Ubicaciones"
            Visible         =   0   'False
         End
         Begin VB.Menu mnu_Listatallas 
            Caption         =   "Lista de tallas"
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
         End
      End
      Begin VB.Menu manra_01 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_salir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu mnulistado 
      Caption         =   "&Listados"
      Begin VB.Menu mnu_catarticulo 
         Caption         =   "Católogo de Artículo"
      End
      Begin VB.Menu mnu_catproveed 
         Caption         =   "Católogo de Proveedores"
      End
   End
   Begin VB.Menu menu03 
      Caption         =   "&Ingresos"
      Begin VB.Menu menu03_01 
         Caption         =   "Nota de Ingreso"
         Shortcut        =   ^I
      End
      Begin VB.Menu menu03_02 
         Caption         =   "Nota de Salidas "
         Shortcut        =   ^S
      End
      Begin VB.Menu menu03_03 
         Caption         =   "Guías de Remisión"
         Shortcut        =   ^G
      End
      Begin VB.Menu menu03_05 
         Caption         =   "Traslado entre Almacenes"
      End
      Begin VB.Menu menu03_08 
         Caption         =   "Emision de Orden de Compra(1)"
      End
      Begin VB.Menu menu03_06 
         Caption         =   "Emison de Orden de compra"
         Enabled         =   0   'False
      End
      Begin VB.Menu menu03_07 
         Caption         =   "Ingreso por Orde. Compras"
      End
   End
   Begin VB.Menu mnucons 
      Caption         =   "&Consultas"
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
      End
      Begin VB.Menu mnu_movarti 
         Caption         =   "Movimiento por Articulo"
      End
   End
   Begin VB.Menu mnurep 
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
         Begin VB.Menu mnu_consxcc_04 
            Caption         =   "Consumo x Centro Costos"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnu_articulocxcosto_04 
            Caption         =   "Consumo Articulos en Centro de Costos"
         End
         Begin VB.Menu mnu_artven_06 
            Caption         =   "Artículos Vencidos"
         End
         Begin VB.Menu mnu_artven_07 
            Caption         =   "Utilidades Venta/Costo"
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
         Caption         =   "Inventario Valorado"
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
         Begin VB.Menu mnu_valxdoc04 
            Caption         =   "Kardex Valorizado por Trans. por Dcmto"
         End
         Begin VB.Menu mnu_valxdoc05 
            Caption         =   "kardex Valorizado por Trans. por Art."
         End
         Begin VB.Menu mnu_valxdoc06 
            Caption         =   "kardex Valorizado por Trans. - Art. Resumido"
         End
      End
      Begin VB.Menu mnu_repdoc 
         Caption         =   "Documentos"
      End
      Begin VB.Menu mnu_Trasla 
         Caption         =   "Informe de Traslados"
      End
   End
   Begin VB.Menu Pro 
      Caption         =   "&Procesos"
      Begin VB.Menu Men_ProVal 
         Caption         =   "Valorización"
         Begin VB.Menu Men_ProCieMen_01 
            Caption         =   "Cierre Mensual"
         End
         Begin VB.Menu mnu_contrans_06_01 
            Caption         =   "Valorización Mensual"
         End
      End
      Begin VB.Menu Men_ProEsp 
         Caption         =   "Especiales"
         Begin VB.Menu Men_EspInvFis_01 
            Caption         =   "Inventario Físico"
            Begin VB.Menu Men_EspInvFis_01_01 
               Caption         =   "Registro de Existencias"
            End
            Begin VB.Menu Men_EspInvFis_01_02 
               Caption         =   "Informes de Inventario Fisico"
            End
         End
         Begin VB.Menu mnu_Asiento_02 
            Caption         =   "Generación de Asiento"
         End
         Begin VB.Menu mnu_RestSalAnt_04 
            Caption         =   "Restaurar Saldo Ant."
         End
         Begin VB.Menu mnu_contrans_05 
            Caption         =   "Conf. de transferencia"
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
      End
      Begin VB.Menu Men_TraVal 
         Caption         =   "Valorización Art. Pendientes"
      End
      Begin VB.Menu Men_TraCor 
         Caption         =   "Correción Art Valorizados"
      End
      Begin VB.Menu mnu_recstk_03 
         Caption         =   "Recalculo Stock"
         Begin VB.Menu mnu_contrans_08 
            Caption         =   "Recalculo Saldo Fisico"
         End
         Begin VB.Menu mnu_recstk_03_01 
            Caption         =   "Recalculo de Stock por Articulos"
         End
         Begin VB.Menu mnu_recstk_03_02 
            Caption         =   "Recalculo de Stock por Series/Lotes"
         End
      End
      Begin VB.Menu mnu_contrans_09 
         Caption         =   "Anulacion de Documentos"
      End
   End
   Begin VB.Menu mnu_auditor01 
      Caption         =   "Auditoria"
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
      End
   End
   Begin VB.Menu sis 
      Caption         =   "Confi&guración"
      Begin VB.Menu Men_SisCrea 
         Caption         =   "Crear"
         Begin VB.Menu mnu_Emp_01 
            Caption         =   "Empresas"
         End
         Begin VB.Menu Men_SisUsu_02 
            Caption         =   "Usuarios"
         End
         Begin VB.Menu Men_SisAdminis_03 
            Caption         =   "Admnistradores"
         End
      End
      Begin VB.Menu Men_SisPar 
         Caption         =   "Parámetros"
      End
      Begin VB.Menu sisra_03 
         Caption         =   "-"
      End
      Begin VB.Menu Men_SisCam 
         Caption         =   "Cambiar Empresa"
         Shortcut        =   {F12}
      End
      Begin VB.Menu Men_SisAl 
         Caption         =   "Cambiar Almacén"
         Shortcut        =   {F11}
      End
   End
End
Attribute VB_Name = "FrmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adoreg As ADODB.Recordset
Dim rs As ADODB.Recordset

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim I As Integer
For I = Forms.count - 1 To 0 Step -1
  Unload Forms(I)
Next

End Sub

Private Sub Men_AlmKar_02_Click()
   FormKardex.Show 1
End Sub

Private Sub Men_AlmStock_01_Click()
    FormStkAlm.Show 1
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
   FrmArFam.Show 1                           ' Form5.show 1
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
  FormModificar.Show 1
End Sub

Private Sub Men_ColorAvios_Click()
FrmColorAvio.Show 1
End Sub

Private Sub Men_DensiTela_19_Click()
FrmDesidadTela.Show 1
End Sub

Private Sub Men_EliDoc_02_Click()
VGElimina = True
FormEliminaDoc.Show 1
End Sub

Private Sub Men_EspInvFis_01_01_Click()
   frmRegistroInventarioFisico.Show 1
   'frmtoma.show 1
End Sub

Private Sub Men_EspInvFis_01_02_Click()
'frmtoma.show 1
frmInformeInventarioFisico.Show 1
End Sub

'Private Sub Men_EspInvFis_01_Click()
'
'End Sub
Private Sub Form_Activate()
  If VGSALIR Then
        Unload Me
  End If
End Sub
Private Sub Form_Load()
 Dim sFileName As String
 Dim sBD As String
 Dim sBDt As String
 Dim n As String
 Dim RSQL As String
 Dim IASA As String
   On Error GoTo Err
   
   'Verifica si es Copia Ilegal
 '  Verificar_Sistema
   VGCodMon = "MN"
   VGtransp = True
   VGSALIR = False
   mensaje1 = "Prueba - Inventarios"
   sFileName = App.Path & "\Inventario.ini"
   sName = sGetIni(sFileName, "CONFIG", "DBNAMES", "?")
   sBD = sGetIni(sFileName, "CONFIG", "BDCONTA", "?")
   sBDt = sGetIni(sFileName, "CONFIG", "BDTRANS", "?")
   cRutP = sGetIni(sFileName, "CONFIG", "RPT", "?")
   VGServer = sGetIni(sFileName, "CONFIG", "SERVER", "?")
   VGBase = sGetIni(sFileName, "CONFIG", "BASE", "?")
   VGBUsuario = sGetIni(sFileName, "CONFIG", "USUARIO", "?")
   VGPassw = sGetIni(sFileName, "CONFIG", "PASS", "?")
   
   
   VGServer2 = sGetIni(sFileName, "CONFIG", "SERVER2", "?")
   VGBase2 = sGetIni(sFileName, "CONFIG", "BASE2", "?")
   VGBUsuario2 = sGetIni(sFileName, "CONFIG", "USUA2", "?")
   VGPassw2 = sGetIni(sFileName, "CONFIG", "PASS2", "?")
   
   VGBase3 = sGetIni(sFileName, "CONFIG", "BASE3", "?")
   
   VGDIRE = sGetIni(sFileName, "CONFIG", "DIRE", "?")
   
   GPunto = sGetIni(sFileName, "CONFIG", "PUNTO", "?")
   
   If sBD <> "" Then VGNameCont = sBD
   If sBDt <> "" Then VGContTra = sBDt           'Base de datos de transacciones
   If sName <> "?" Then
        RUTA = sName '(?)
        If cRutP = "?" Then
           cRutP = RUTA & "Reportes\"
        End If
        VGLongCodigo = 8
        'El tipo de almacen
        central FrmPrincipal
        Set cConexCom = New ADODB.Connection  'BD. Común
        Set cConexCont = New ADODB.Connection  'BD. Contabilidad
        'frmInicio.show 1
         frmlogin.Show 1
        FrmPrincipal.Caption = "Sistema de Inventario" & "     " & VGNomAlm & "    " & VGNemp
        If VGSALIR Then
            If cConexCom.State = 1 Then cConexCom.Close
            If cConexCont.State = 1 Then cConexCont.Close
            FrmPrincipal.Visible = False
            Form_Unload (0)
            Exit Sub
        End If
    Else
        MsgBox "No encontró la BD", vbCritical, mensaje1
        FrmPrincipal.Visible = False
        Form_Unload (0)
        Exit Sub
    End If
    VGAutomatico = False
    FrmPrincipal.mnu_ajuste.Visible = False
    FrmPrincipal.mnu_artven_06.Visible = False
    'FrmPrincipal.mnu_repIASA.Visible = False
    'FrmPrincipal.mnu_guiaIngIasa.Visible = False
    'FrmPrincipal.mnu_recepcion.Visible = False
    FrmPrincipal.mnu_Asiento_02.Visible = True
    
    'FrmPrincipal.mnu_repIASA.Visible = True
    'FrmPrincipal.mnu_guiaIngIasa.Visible = True
    'FrmPrincipal.mnu_recepcion.Visible = True

    
    RSQL = "select * from configuracion"
    If cConexCom <> "" Then
            Set adoreg = New ADODB.Recordset
            adoreg.Open RSQL, cConexCom, adOpenDynamic, adLockOptimistic
            If Not adoreg.EOF Then
                    IASA = IIf(IsNull(adoreg("cod_iasa")), "", adoreg("cod_iasa"))
'                    If IASA <> "" Then
'                      FrmPrincipal.mnu_repIASA.Visible = True
'                      FrmPrincipal.mnu_guiaIngIasa.Visible = True
'                     ' FrmPrincipal.mnu_recepcion.Visible = True
'                   End If
                     If adoreg("cod_bloqueo") Then
                          ' variable global par el setear
                           VGAutomatico = True
                     End If
            End If
            adoreg.Close
   End If
   Exit Sub
Err:
    MsgBox Err.Description, vbExclamation, "Aviso"
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
FormGuiaSal.Show 1
End Sub

Private Sub Men_GuiEli_01_Click()
VGElimina = False
FormEliminaDoc.Show 1
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
  FrmProvee.Show 1
End Sub

Private Sub Men_ManTra_Click()
  FormTransa.Show 1
End Sub

Private Sub Men_MedidaAvios_Click()
FrmMedidaAvio.Show 1
End Sub

Private Sub Men_MezclaTela_17_Click()
FrmMezclaTela.Show 1
End Sub

Private Sub Men_mnGui_Click()
'If Not VGLadrillera Then
  VGGuiaSal = True
  VGRegEnt = 2
  FormGuiaSal.Show 1
  
'Else
'  VGGuiaSal = True
'  FormGuiaSalLadrillo.show 1
'End If
End Sub

Private Sub Men_mnu_alma_Click()
  FormAlmacen.Show 1
End Sub

Private Sub Men_TituTela_16_Click()
FrmTituloTela.Show 1
End Sub

Private Sub menu03_01_Click()
   VGRegEnt = 1
   FormRegistro.Show 1
End Sub

Private Sub menu03_02_Click()
   VGRegEnt = 0
   FormRegistro.Show 1
End Sub

Private Sub menu03_03_Click()
  VGGuiaSal = True
  VGRegEnt = 2
  FormGuiaSal.Show 1
End Sub

Private Sub menu03_05_Click()
FrmTraslado.Show
End Sub

Private Sub menu03_06_Click()
    frmEmisionOC.Show 1
End Sub

Private Sub menu03_07_Click()
    frmingresoOC.Show
End Sub

Private Sub menu03_08_Click()
frmTraEmi.Show
End Sub

Private Sub mnu_articulocxcosto_04_Click()
frmArticuloXCenCos.Show 1
End Sub

Private Sub mnu_artven_06_Click()
'   MsgBox "No existe aticulos vencidos", vbInformation, mensaje1
  FrmArtVen.Show 1
End Sub

Private Sub mnu_artxven_Click()
 ' MsgBox "No existe aticulos por vencer", vbInformation, mensaje1
   FrmArtVen.Show 1
End Sub

Private Sub mnu_artven_07_Click()
        frmUtilidadVentaCosto.Show 1
End Sub

Private Sub mnu_artven_08_Click()
   FrmStockInicial.Show 1
End Sub

Private Sub mnu_Asiento_02_Click()
  'FrmAsiento.show 1
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
   'frmCenCos.Show 1
End Sub

Private Sub mnu_contrans_05_Click()
 'FrmConfdeTrans.show 1
 FrmConfdeTrans2.Show 1
End Sub

Private Sub mnu_contrans_06_01_Click()
  FrmKardexRevalorizaMes.Show 1
End Sub

Private Sub mnu_contrans_06_03_Click()
   FrmKardexRevalorizaMes.Show 1
End Sub

Private Sub mnu_contrans_08_Click()
  PrcSaldos.Show 1
End Sub

Private Sub mnu_contrans_09_Click()
 PrcAnulaDocumento.Show 1
End Sub

Private Sub mnu_conValArtPend_Click()
 FormConValArt.Show 1
End Sub

Private Sub mnu_defdoc_04_Click()
 FrmCfgDocumento.Show 1
End Sub



Private Sub mnu_Distrito_09_Click()
FrmArTabAyu.G_nOpc = 7
FrmArTabAyu.G_cTabla = "13"
FrmArTabAyu.Show 1
End Sub

Private Sub mnu_docum_01_Click()
FrmArDocumento.Show 1
End Sub

Private Sub mnu_docvalorizado_Click()
FormConDocVal.Show 1
End Sub

Private Sub mnu_Emp_01_Click()
 FrmCfgEmpresa.Show 1
End Sub

Private Sub mnu_giropro_07_Click()
  FrmArTabAyu.G_nOpc = 10
  FrmArTabAyu.G_cTabla = "62"
  FrmArTabAyu.Show 1
End Sub

Private Sub mnu_grupos_Click()
'    FrmArGrupos.show 1
End Sub
Private Sub mnu_IngXordCompra_Click()
    frmingresoOC.Show 1
'    frmTraEmi.show 1
End Sub

Private Sub mnu_kdxcencos_02_Click()
  FormKarValC.Show 1
End Sub

Private Sub mnu_Lineas_Click()
'  FrmArLineas.show 1
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

Private Sub mnu_recepcion_Click()
Dim RSQL As String
    RSQL = "select conf_codigo from configuracion"
    Set rs = New ADODB.Recordset
    rs.Open RSQL, cConexCom, adOpenStatic
    If rs.EOF Then
        MsgBox "No se ha definido el codigo de IASA", vbExclamation, "Aviso"
        rs.Close
        Exit Sub
    End If
    rs.Close
    FrmRegistrarIASA.Show 1
End Sub

Private Sub mnu_recstk_03_01_Click()
' FrmActSal.show 1
 '*********Mofifcado Por Roberto Maza Milla  06/07/2001
 frmRestaurarSaldos.Show 1
  
End Sub

Private Sub mnu_recstk_03_02_Click()
frmCalcularLote_Serie2.Show 1
End Sub



Private Sub mnu_repdoc_Click()
    'frmrepdoc.show 1
    RepDocuDeta.Show 1
End Sub

Private Sub mnu_repdocAlm_Click()
   FrmRepMovFec.Show 1
End Sub





Private Sub mnu_RestSalAnt_04_Click()
 frmRestSalAnt.Show 1
End Sub



Private Sub mnu_salir_Click()
Unload Me
End Sub

Private Sub mnu_sincroTC_Click()
        FrmSincronizaTC.Show 1
''
''Dim SUBICA As String
''Dim SQL As String
''
''    If UCase(Dir$(cRuta4)) = UCase(VGNameCont & ".MDB") Then
''       SUBICA = "[" & cRuta4 & "].TIPO_CAMBIO T"
''    Else
''       If UCase(Dir$(cRuta2)) = "BDCOMUN.MDB" Then
''          SUBICA = "[" & cRuta2 & "].TIPO_CAMBIO T"
''       End If
''    End If
''
''SQL = "Update ( MOVALMDET  D INNER JOIN MOVALMCAB C ON D.DEALMA=C.CAALMA AND D.DETD=C.CATD AND D.DENUMDOC=C.CANUMDOC) "
''SQL = SQL & "INNER JOIN  " & SUBICA & " ON C.CAFECDOC=T.TIPOCAMB_FECHA  "
''SQL = SQL & " SET C.CATIPCAM=T.TIPOCAMB_VENTA,D.DETIPCAM=T.TIPOCAMB_VENTA "
''
''cConexCom.Execute SQL

'SASA
End Sub

Private Sub mnu_stkArt1_Click()
Dim RSQL As String
'Dim db As Database
Dim rs As Recordset
If Trim(VGAlma) <> "" Then
  'RSQL = "select  stcodigo from  StkArt  where  STALMA = '" & VGAlma & "' "
  RSQL = "select  stcodigo from  StkArt "
  'Set db = Workspaces(0).OpenDatabase(cRuta2)
  Set rs = cConexCom.Execute(RSQL) '  db.OpenRecordset(RSQL, dbOpenSnapshot)
  If rs.RecordCount = 0 Then
           MsgBox "No hay articulos en este almacen", vbCritical, mensaje1
  Else
           FormConStk.Show 1
  End If
  rs.Close
 ' db.Close
End If
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
  RepTraslados.Show 1
End Sub

Private Sub mnu_Traslado_Click()
FrmTraslado.Show 1
End Sub

Private Sub mnu_ubica13_Click()
FrmMantenUbica.Show 1
End Sub

Private Sub mnu_unidades_02_Click()
 ' FrmArUniMed.show 1
 FrmArUnidades.Show 1
End Sub

Private Sub mnu_valxdoc_03_Click()
 ' frmkardexDoc.show 1
 VGRepKxVal = 0
 'FormKardexValTXDocumento.Show 1
 RepKardexValTXDocumento.Show 1
End Sub

Private Sub Men_mnulogistica_Click()
   Formlogistica.Show 1
End Sub

Private Sub Men_ProCieMen_01_Click()
Dim RSQL As String
Dim Rsql1 As String
Dim rs1 As ADODB.Recordset
Dim rs As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim nMes As Integer
Dim nAnno As Integer
Dim nTra As Integer
Dim Ado As ADODB.Recordset
Dim cMesPro As String
Dim nPrePro As Double
Dim nTipCam As Double

On Error GoTo ErrProCierre

If MsgBox("Desea realizar el Cierre Mensual", vbQuestion + vbYesNo, "Mensaje") = vbNo Then Exit Sub

Rsql1 = "Select n.STCODIGO FROM  StkArt n where n.STALMA = '" & VGAlma & "'  "
Set rs1 = New ADODB.Recordset
rs1.Open Rsql1, cConexCom, adOpenStatic
If rs1.RecordCount = 0 Then
      MsgBox "No hay articulos en el respectivo Almacen", vbInformation, mensaje1
      rs1.Close
      Exit Sub
End If
rs1.Close
'Indica mes del ultimo cierre
nMes = Month(Date)
nAnno = Year(Date)
RSQL = "SELECT TOP 1 min(CAFECDOC) From MovAlmCab  WHERE isnull(CACIERRE,0) = 0 AND CAALMA = '" & VGAlma & "'"
Set Rs2 = New ADODB.Recordset
Rs2.Open RSQL, cConexCom, adOpenStatic
If Not Rs2.EOF Then
    Rs2.MoveFirst
    If IsNull(Rs2(0)) Then
       MsgBox "No hay Informacion que tenga Cierre Pendiente ", vbInformation, "Aviso"
       Exit Sub
    End If
       nMes = Month(Rs2(0))
       nAnno = Year(Rs2(0))
    If nMes = 13 Then
         nMes = 1
    End If
    MsgBox "Se va a ralizar el cierre del  mes : " & MonthName(nMes), vbInformation, "Aviso"
Else
    MsgBox "No hay Informacion que tenga Cierre Pendiente ", vbInformation, "Aviso"
    Exit Sub
End If

Screen.MousePointer = 1

If MsgBox("Está Ud. seguro de realizar el Cierre Mensual del mes " & MonthName(nMes), vbQuestion + vbYesNo, "Mensaje") = vbNo Then Exit Sub
Screen.MousePointer = 11
'Verifica si seha valorizado los Articulos
RSQL = "Select  p.ACODIGO, p.ADESCRI, m.CACODMOV ,n.DENUMDOC, m.CANOMPRO, m.CARFTDOC " & _
          "from MaeArt p,MovAlmCab m, MovAlmDet n   " & _
          "where  m.CAALMA ='" & VGAlma & _
          "' AND  n.DEALMA = m.CAALMA and CATIPMOV='I'  and p.ACODIGO = n.DECODIGO and  m.CASITGUI <> 'A'  and " & _
          " n.DEPRECIO = 0  And  n.DENUMDOC  = m.CANUMDOC and  n.DETD = m.CATD   AND MONTH(CAFECDOC) <= " & nMes & " AND YEAR(CAFECDOC) = " & nAnno & " ORDER BY m.CANUMDOC"

Set rs = New ADODB.Recordset
rs.Open RSQL, cConexCom, adOpenStatic
If rs.RecordCount > 0 Then
   MsgBox "Debe Valorizar todos sus Articulos Pendientes", vbInformation, mensaje1
   rs.Close
   Screen.MousePointer = 1
   Exit Sub
End If
rs.Close

'Consulta de todos los articulos valorizados
RSQL = "Select  p.ACODIGO, p.ADESCRI, m.CACODMOV ,n.DENUMDOC, m.CANOMPRO, m.CARFTDOC " & _
          "from MaeArt p,MovAlmCab m, MovAlmDet n   " & _
          "where  m.CAALMA ='" & VGAlma & _
          "' AND  n.DEALMA = m.CAALMA and CATIPMOV='I'  and p.ACODIGO = n.DECODIGO and  m.CasitGUI <> 'A'  and " & _
          " n.DEPRECIO <> 0  And m.CANUMDOC= n.DENUMDOC AND   MONTH(CAFECDOC) <= " & nMes & " AND YEAR(CAFECDOC) = " & nAnno & " ORDER BY m.CANUMDOC"

Set rs = New ADODB.Recordset
rs.Open RSQL, cConexCom, adOpenStatic

RSQL = "Select * from MovAlmCab Where isnull(CACIERRE,0) = 0  AND CAALMA = '" & VGAlma & "' AND MONTH(CAFECDOC) = " & nMes & " AND YEAR(CAFECDOC) = " & nAnno
Set Ado = New ADODB.Recordset
Ado.Open RSQL, cConexCom, adOpenStatic
If Ado.RecordCount > 0 Then
    Rsql1 = "Select * FROM MovAlmCab Where isnull(CACIERRE,0) = 0 AND CAALMA = '" & VGAlma & "' AND MONTH(CAFECDOC) < " & nMes & " AND YEAR(CAFECDOC) = " & nAnno & " "
    Set rs1 = New ADODB.Recordset
    rs1.Open Rsql1, cConexCom, adOpenStatic
    If rs1.RecordCount > 0 Then
        MsgBox "No se ha hecho el Proceso de Cierre Mensual en los meses anteriores a este  Mes", vbInformation, "Verificar"
        rs1.Close: Ado.Close: rs.Close
        Screen.MousePointer = 1
        Exit Sub
    End If
    rs1.Close
    'Aqui se hace el cierre mensual
    Rsql1 = "Update MovAlmCab set CACIERRE =  1 " & _
                " where  CAALMA = '" & VGAlma & "'   AND MONTH(CAFECDOC) =" & nMes & " AND YEAR(CAFECDOC) = " & nAnno
    cConexCom.Execute Rsql1
    nTra = 1
    cConexCom.BeginTrans
    cConexCom.Execute Rsql1
    cConexCom.CommitTrans
    nTra = 0
    If rs.EOF Then
      MsgBox "No hay registro para cerrar para el mes respectivo", vbInformation, "Inventarios"
      rs.Close
      Screen.MousePointer = 1
      Exit Sub
    Else
      rs.MoveFirst
    End If
'RMM**************************************************
'DESHABILITADO POR ROBERTO MAZA EL COSTO PROMEDIO NO PUEDE SER EL MISMO PARA CADA MES QUE CIERRO
'CASO QUE RECIEN EN MARZO EMPIEZO HA CERRAR EL MES DE ENERO, ACTUALIZARA EL MORESMES CON EL COSTO PROMEDIO ACTUAL
'RMM**************************************************
'    Do While Not rS.EOF
'        cMesPro = nAnno & Format(nMes, "00")
'        nPrePro = Val(Devolver_Dato(1, rS("acodigo"), "StkArt", "STCODIGO", False, "STKPREPRO", VGAlma, "STALMA"))
'
'        If UCase(Dir$(cRuta4)) = UCase(cNomBd4) Then
'            nTipCam = Val(Devolver_Dato(3, Date, "Tipo_Cambio", "TIPOCAMB_FECHA", True, "TIPOCAMB_VENTA", "ME", "TIPOMON_CODIGO"))
'        ElseIf UCase(Dir$(cRuta2)) = "BDCOMUN.MDB" Then
'            nTipCam = Val(Devolver_Dato(1, Date, "Tipo_Cambio", "TIPOCAMB_FECHA", True, "TIPOCAMB_VENTA", "ME", "TIPOMON_CODIGO"))
'        End If
'
'        Rsql1 = "Update  MoResMes  Set  SMMNPREUNI  = " & nPrePro & ", SMUSPREUNI  = " & nPrePro / IIf(nTipCam = 0, 1, nTipCam) & "  where  SMALMA = '" & VGAlma & "' and SMMESPRO = '" & cMesPro & "' AND  SMCODIGO= '" & rS("acodigo") & "'"
'        nTra = 1
'        cConexCom.BeginTrans
'        cConexCom.Execute Rsql1
'        cConexCom.CommitTrans
'        nTra = 0
'
'        rS.MoveNext
'        If rS.EOF Then Exit Do
'    Loop
'**************************************************
    rs.Close
    MsgBox "Se Finalizó el Cierre Mensual", vbInformation, "Información"
Else
    MsgBox "Ya se realizó el Cierre Mensual de este Mes " & MonthName(nMes) & Chr(13) & "o No existe ningún movimiento en este Mes", vbInformation, "Información"
    Screen.MousePointer = 1
    Exit Sub
End If
Ado.Close
Screen.MousePointer = 1

cMesPro = nAnno & Format(nMes, "00")
cConexCom.BeginTrans
cConexCom.Execute "delete from CIERRMESVALOR WHERE CIERRMES='" & cMesPro & "' AND Cierralma='" & VGAlma & "'"
cConexCom.Execute "INSERT INTO  CIERRMESVALOR (CierrMes,CierrFech,CierrOper,Cierralma)VALUES('" & cMesPro & "'," & Format(Now, "dd/mm/yyyy") & ",'RMAZA','" & VGAlma & "')"
cConexCom.CommitTrans

Exit Sub
ErrProCierre:
    MsgBox Err.Description
    If nTra = 1 Then cConexCom.RollbackTrans
    nTra = 0
    Screen.MousePointer = 1
End Sub
'AND MONTH(CAFECDOC) <= " & nMes & " AND YEAR(CAFECDOC) <= " & Year(Date) & "
Private Sub Men_SisAdminis_03_Click()
FrmCfgAdminist.Show 1
End Sub

Private Sub Men_SisAl_Click()
  FormCamAlm.Show 1
End Sub

Private Sub Men_SisCam_Click()
FrmCfgCambioEmp.Show 1
End Sub

Private Sub Men_SisPar_Click()
    FormConfiguracion.Show 1
End Sub

Private Sub Men_SisUsu_02_Click()
'  frmUsuario.show 1
  frmCfgUsuario.Show 1
End Sub

Private Sub Men_TraCor_Click()
  FormCorrArt.Show 1
End Sub

Private Sub Men_TraRegEnt_Click()
   VGRegEnt = 1
   FormRegistro.Show 1
End Sub

Private Sub Men_TraRegSal_Click()
  VGRegEnt = 0
  FormRegistro.Show 1
End Sub

Private Sub Men_TraVal_Click()
    FormValArtPed.Show 1
End Sub

Private Sub mnu_valxdoc04_Click()
    VGRepKxVal = 1
    FormKardexValTDoc.Show 1
End Sub

Private Sub mnu_valxdoc05_Click()
  VGRepKxVal = 2
  FormKardexValTDoc.Show 1
End Sub

Private Sub mnu_valxdoc06_Click()
  VGRepKxVal = 3
  FormKardexValTDoc.Show 1
End Sub
