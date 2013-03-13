Attribute VB_Name = "MainCobrar"
Public Cadenabusca As String

Public cn As New ADODB.Connection

Public VGformatofecha As String
Public VGPlanillaAjuste As Integer

Public VGtipo As TIPOSISTEMA


'Variables de Ventana de Ayuda
Public nAyuda As String
Public nDetalle As String
Public nfecha As String

Public g_BaseContab As String
Public g_Empresa As String
Public g_DetalleEmpresa As String
Public g_PedidoPuntoVta As String
Public g_DetallePuntoVta As String

Public VGvardllgen As dllgeneral.dll_general 'Dll de Algunas funciones

Public g_TipoMovi As String
Public g_usuario As String
Public g_ptoventa As String
Public g_pedserie As String
Public g_facserie As String
Public g_bolserie As String
Public g_guiaserie As String

Public punto As PuntoVenta
Public Tipodocu As Tipodocu

'Variables de acceso de usuario
Public Const g_tipoped = "PE"    'PEDIDO
Public Const g_tipofac = "01"    'factura
Public Const g_tipobol = "03"    'boletas
Public Const g_tipoguia = "80"    '9"   'guias B.O
Public Const g_Todos = "99"      'Todos los documentos para Reporte
Public Const g_Eventual = "88888888888" 'cliente eventual

'REPORTES

Public RutaRep  As String
Public RutaRepProc As String

Public VGparametros As Parametrocobrar
Public VGparamsistem As PuntoVenta
Public modoventa As modoventa

Private Type Parametrocobrar         ' Crea Tipo de Empresa
   nombre As String * 2
   moneda As String * 2
   tieneigv As String * 1
   igv As Double
   tienedscto As String * 1
   descuento As Double
   empresacodigo As String * 2
   mensaje As String * 70
   almacen As String * 2
   listapre As String * 1
   tipocambio As Double
   comivende As String * 1
   formaemi As String * 1
   paraboleta As String * 1
   NomEmpresa As String
   RucEmpresa As String
   sistemamultiempresas As Boolean
   contabilizaenlinea As Boolean
   imprimevoucher As Integer
   puntovta As String
   
   
End Type


Private Type PuntoVenta   'Crea Punto de Venta

    nropedido As String * 1
    nroguia As String * 1
    nrofactura As String * 1
    nroboleta As String * 1
    nroabono As String * 1
    nrocargo As String * 1
    ventaauto As String * 1
    seriefac As String * 3
    seriebol As String * 3
    serieguia As String * 3
    serieped As String * 3
    serie99 As String * 3
    
    Anoproceso As String * 4
    puntovta As String * 2
    
    bdempresa As String
    Usuario As String
    pwd As String
    servidor As String
    
    bdempresaCT As String
    usuarioCT As String
    pwdCT As String
    servidorCT As String
    
    bdempresaGEN As String
    usuarioGEN As String
    pwdGEN As String
    servidorGEN As String
    
    RutaReport As String
    carpetareportes As String
End Type

Private Type modoventa   'Crea modoventa
    descuento As String
    impuestos As String
    nroitem  As Double
    numeraauto As String
    ctrlinventario As String
    unidadmedida As String
    copiasfac As Double
    copiasbol As Double
    copiashoja As Double
    ctacte As String * 1
    ingcliente As String
    ingforma As String
    emitehoja As String
    emiteguia As String
    emitefact As String
    modificaguia As String
    ingpedido As String
    inghoja As String
    ingguia As String
    usafactor As String
    documento As String

End Type


Private Type Tipodocu   'Tipo de documento
    numeauto As String
    numerador As String
End Type

