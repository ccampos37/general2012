VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form UpdateDatabases 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Verificación de Datos de la Empresa"
   ClientHeight    =   2520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7215
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   7215
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Salir"
      Height          =   375
      Left            =   3996
      TabIndex        =   4
      Top             =   1836
      Width           =   1575
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   228
      Left            =   1908
      TabIndex        =   3
      Top             =   1476
      Width           =   4908
      _ExtentX        =   8652
      _ExtentY        =   423
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Timer Timer1 
      Left            =   4320
      Top             =   1920
   End
   Begin VB.PictureBox PictureClip2 
      Height          =   384
      Left            =   120
      ScaleHeight     =   330
      ScaleWidth      =   900
      TabIndex        =   2
      Top             =   2640
      Width           =   960
   End
   Begin VB.CommandButton CmdInit 
      Caption         =   "Continuar"
      Height          =   375
      Left            =   1980
      TabIndex        =   1
      Top             =   1836
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   $"Bdchecks.frx":0000
      Height          =   732
      Left            =   1944
      TabIndex        =   5
      Top             =   288
      Width           =   4872
   End
   Begin VB.Image Image1 
      Height          =   2076
      Left            =   180
      Picture         =   "Bdchecks.frx":00A5
      Stretch         =   -1  'True
      Top             =   216
      Width           =   1500
   End
   Begin VB.Label LblMsg 
      Height          =   252
      Left            =   1908
      TabIndex        =   0
      Top             =   1140
      Width           =   4920
   End
End
Attribute VB_Name = "UpdateDatabases"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Ejecutando As Boolean
Dim WithEvents cConexAux As adodb.Connection
Attribute cConexAux.VB_VarHelpID = -1
Dim Y As Integer
Dim X As Integer
Dim DELAY As Integer
Dim toggle As Integer
Dim flap As Integer
Private Sub runtop()
    ' Avanza la animación una imagen.
    Y = Y + 1: If Y = 18 Then Y = 0
    ' Animación de icono:  sólo se verá cuando el formulario esté
    ' minimizado. En la matriz Picture3 se han cargado archivos de
    ' icono (.ICO), y no .BMP. Esto permite usar la función de
    ' máscara del archivo de icono, dejando que el fondo que hay
    ' tras el icono se muestre a través de él.
   'UpdateDataBases.Icon = Image1(Y).Picture
End Sub

Private Sub CmdGirar_Click()
 Timer1.Interval = 1
 Timer1.Enabled = True
 If toggle = 1 Then
   toggle = 0
 Else
   toggle = 1
 End If
 'runtop
End Sub

Private Sub CmdInit_Click()
 If Not Ejecutando Then
    'Picture6.Visible = True
    toggle = 1
    runtop
    Timer1.Interval = 1
    Timer1.Enabled = True
    Ejecutando = True 'Para que no haga caso
    CmdInit.Visible = False
    UpdateDatabases.Caption = "Verificando Información de la Empresa"
    'Set Conex = New ADODB.Connection
    'Conex.ConnectionString = cConexCom.ConnectionString
    'Conex.CursorLocation = cConexCom.CursorLocation
    'Conex.Provider = cConexCom.Provider
    'Conex.Open
    ActualizarBD2
    'Conex.Close
    Set Conex = Nothing
    Ejecutando = False
    LblMsg = "Ingresando..."
    Me.Refresh
    UpdateDatabases.Caption = "Verificación Concluída. Ingresando..."
    toggle = 0
    Unload Me
 End If
End Sub

Public Sub ActualizarBD2()
Dim SQL As String
On Local Error GoTo ERRAR
 Screen.MousePointer = 11
'*--------------------------------------*
 ProgressBar1.Min = 1
 ProgressBar1.Max = 80
 ProgressBar1.Visible = True
 ProgressBar1.Value = 1
 LblMsg.Visible = True
 LblMsg = "Verificando Base de Datos..."
 Me.Refresh
 
 ProgressBar1.Value = ProgressBar1.Value + 1
 DoEvents
 
 LblMsg = "Verificando Tabla de Registro de Kits..."
 ProgressBar1.Value = ProgressBar1.Value + 1
 DoEvents
 
  If Not ExisteElem(0, cConexCom, "KITS") Then
      SQL = " Create Table KITS (CODART Text(20),CODKIT Text(20), " & _
      "  CANART double)"
      cConexCom.Execute SQL
  End If
  
 LblMsg = "Verificando Tabla de Maestro de Articulos..."
 ProgressBar1.Value = ProgressBar1.Value + 1
 DoEvents
  
 If Not ExisteElem(1, cConexCom, "MAEART", "AMARCA") Then
    cConexCom.Execute "ALTER TABLE MAEART ADD COLUMN   AMARCA  TEXT(20)" '
 End If
 
 ProgressBar1.Value = ProgressBar1.Value + 1
 DoEvents
  
 If Not ExisteElem(1, cConexCom, "MAEART", "ACOLOR") Then
    cConexCom.Execute "ALTER TABLE  MAEART  ADD COLUMN   ACOLOR  TEXT(20)" '
 End If
 
 '*****************************************************************
 '*** ULTIMA ACTUALIZACION 28/06/2001    ROBERTO M.M.
 '*****************************************************************
 LblMsg = "Verificando Tabla de Cierre de Mes ..."
 ProgressBar1.Value = ProgressBar1.Value + 1
 DoEvents
   
 If Not ExisteElem(0, cConexCom, "CIERRMESVALOR") Then
       SQL = " Create Table CIERRMESVALOR (CIERRMES Text(6),CIERRALMA Text(2),CIERRFECH DATETIME, CIERROPER TEXT(15) , " & _
       " CONSTRAINT Clave PRIMARY KEY (CIERRMES))"
       cConexCom.Execute SQL
       ProgressBar1.Value = ProgressBar1.Value + 1
       DoEvents
 Else
      If Not ExisteElem(1, cConexCom, "CIERRMESVALOR", "CIERRALMA") Then
          cConexCom.Execute "ALTER TABLE  CIERRMESVALOR  ADD COLUMN  CIERRALMA text(2) " '
      End If
       ProgressBar1.Value = ProgressBar1.Value + 1
       DoEvents
 End If
  
 LblMsg = "Verificando Tabla de Saldos Mensuales ..."
 ProgressBar1.Value = ProgressBar1.Value + 1
 DoEvents
  
 If Not ExisteElem(1, cConexCom, "MORESMES", "SMSALDOINI") Then
    cConexCom.Execute "ALTER TABLE  MORESMES  ADD COLUMN  SMSALDOINI  DOUBLE " '
 End If
   
 LblMsg = "Verificando Tablas de Temporales  ..."
 ProgressBar1.Value = ProgressBar1.Value + 1
 DoEvents
   
 If Not ExisteElem(0, cConexCom, "COSPROFECH") Then
    SQL = " Create Table COSPROFECH ( AUXALMA Text(3),AUXTD Text(3),AUXNUMDOC Text(10),AUXCODART Text(20) ,AUXFECDOC DATETIME,AUXCANT DOUBLE,AUXPRECIO DOUBLE,AUXPRECOS DOUBLE   )" '(AUXTD , AUXNUMDOC , AUXCODART , AUXFECDOC )
    cConexCom.Execute SQL
 End If
  
 ProgressBar1.Value = ProgressBar1.Value + 1
 DoEvents
  
 If Not ExisteElem(1, cConexCom, "KARDEXAUX", "TIPDOCRF") Then
    cConexCom.Execute "ALTER TABLE  KARDEXAUX  ADD COLUMN  TIPDOCRF text(2) " '
 End If
  
 ProgressBar1.Value = ProgressBar1.Value + 1
 DoEvents
  
 If Not ExisteElem(1, cConexCom, "KARDEXAUX", "NUMDOCRF") Then
    cConexCom.Execute "ALTER TABLE  KARDEXAUX  ADD COLUMN  NUMDOCRF text(10) " '
 End If

 ProgressBar1.Value = ProgressBar1.Value + 1
 DoEvents

 If Not ExisteElem(1, cConexCom, "KARDEXAUX", "NOMREFE") Then
    cConexCom.Execute "ALTER TABLE  KARDEXAUX  ADD COLUMN  NOMREFE text(50) " '
 End If
  
 ProgressBar1.Value = ProgressBar1.Value + 1
 DoEvents
  
 Call ADOConectar
  
 If Not ExisteElem(1, cConexAux, "KARDEX_VAL", "ING_SAL") Then
    cConexAux.Execute "ALTER TABLE  KARDEX_VAL  ADD COLUMN  ING_SAL TEXT(20) " '
 End If
 
 cConexAux.Close
 
 'RMM******CASO CHEMEX************************************************************
 LblMsg = "Verificando Tabla de DOCUMENTOS ..."
 ProgressBar1.Value = ProgressBar1.Value + 1
 DoEvents
 
 If Not ExisteElem(1, cConexCom, "NUM_DOCUMENTOS", "CTIMPRESORA") Then
    cConexCom.Execute "ALTER TABLE  NUM_DOCUMENTOS  ADD COLUMN  CTIMPRESORA  text(30)"
 End If
 
 ProgressBar1.Value = ProgressBar1.Value + 1
 DoEvents
 
 If Not ExisteElem(1, cConexCom, "NUM_DOCUMENTOS", "CTPTO") Then
    cConexCom.Execute "ALTER TABLE  NUM_DOCUMENTOS  ADD COLUMN  CTPTO text(1)  " '
 End If
 
 ProgressBar1.Value = ProgressBar1.Value + 1
 DoEvents
 
 If Not ExisteElem(1, cConexCom, "NUM_DOCUMENTOS", "CTSERNUM") Then
    cConexCom.Execute "ALTER TABLE  NUM_DOCUMENTOS  ADD COLUMN  CTSERNUM text(1)  " '
 End If
 
 ProgressBar1.Value = ProgressBar1.Value + 1
 DoEvents
 
 If Not ExisteElem(1, cConexCom, "NUM_DOCUMENTOS", "CTCAMBIO") Then
    cConexCom.Execute "ALTER TABLE  NUM_DOCUMENTOS  ADD COLUMN  CTCAMBIO text(1)  " '
 End If
 
 ProgressBar1.Value = ProgressBar1.Value + 1
 DoEvents
 
 If Not ExisteElem(1, cConexCom, "NUM_DOCUMENTOS", "CTSTOCK") Then
    cConexCom.Execute "ALTER TABLE  NUM_DOCUMENTOS  ADD COLUMN  CTSTOCK text(1)  " '
 End If
 
 ProgressBar1.Value = ProgressBar1.Value + 1
 DoEvents
 
 If Not ExisteElem(1, cConexCom, "NUM_DOCUMENTOS", "CTCONTROLADOR") Then
    cConexCom.Execute "ALTER TABLE  NUM_DOCUMENTOS  ADD COLUMN  CTCONTROLADOR text(30)  " '
 End If
 
 ProgressBar1.Value = ProgressBar1.Value + 1
 DoEvents
 
 If Not ExisteElem(1, cConexCom, "NUM_DOCUMENTOS", "CTPUERTO") Then
    cConexCom.Execute "ALTER TABLE  NUM_DOCUMENTOS  ADD COLUMN  CTPUERTO text(10)  " '
 End If
 
 'RMM******CASO CHEMEX************************************************************
 
 LblMsg = "Verificando Tablas de Invetario Fisico  ..."
 ProgressBar1.Value = ProgressBar1.Value + 1
 DoEvents
 
 If Not ExisteElem(0, cConexCom, "InveFisiCab") Then
       SQL = " Create Table InveFisiCab ( AUXNUMINVE TEXT(10), AUXALMA Text(3),AUXFECH DATETIME ,AUXRESPON TEXT(15),AUXOBSER TEXT(255)" & _
       ", CONSTRAINT Clave PRIMARY KEY ( AUXALMA,AUXNUMINVE )  )"
       cConexCom.Execute SQL
 Else
      If Not ExisteElem(1, cConexCom, "InveFisiCab", "AUXESTADO") Then
         cConexCom.Execute "ALTER TABLE  InveFisiCab  ADD COLUMN  AUXESTADO TEXT(2) " '*****INDICA SI YA SE INGRESO O ESTA PENDIENTE EN INVENTARIO FISICO
      End If
 End If

 ProgressBar1.Value = ProgressBar1.Value + 1
 DoEvents
 
 If Not ExisteElem(0, cConexCom, "InveFisiDet") Then
       SQL = " Create Table InveFisiDet ( AUXNUMINVE TEXT(10), AUXALMA Text(3), AUXFAMIL Text(8),AUXCODART Text(20) ,AUXSTOCK DOUBLE,AUXINGR DOUBLE,AUXDIFE DOUBLE " & _
       ", CONSTRAINT Clave PRIMARY KEY ( AUXALMA , AUXNUMINVE,AUXCODART )  )"
       cConexCom.Execute SQL
 Else
      If Not ExisteElem(1, cConexCom, "InveFisiDet", "AUXFAMIL") Then
         cConexCom.Execute "ALTER TABLE  InveFisiDet  ADD COLUMN  AUXFAMIL Text(8) " '*****INDICA SI YA SE INGRESO O ESTA PENDIENTE EN INVENTARIO FISICO
      End If
       
 End If

 LblMsg = "Verificando Tablas de Configuración  ..."
 ProgressBar1.Value = ProgressBar1.Value + 1
 DoEvents

 If Not ExisteElem(1, cConexCom, "CONFIGURACION", "conf_codigoIng") Then
    cConexCom.Execute "ALTER TABLE CONFIGURACION  ADD COLUMN  conf_codigoIng Text(8) "
 End If
 
 ProgressBar1.Value = ProgressBar1.Value + 1
 DoEvents
 
 If Not ExisteElem(1, cConexCom, "CONFIGURACION", "cosven_debe") Then
    cConexCom.Execute "ALTER TABLE CONFIGURACION  ADD COLUMN  cosven_debe Text(10) "
 End If
 
 ProgressBar1.Value = ProgressBar1.Value + 1
 DoEvents
 
 If Not ExisteElem(1, cConexCom, "CONFIGURACION", "cosven_Habe") Then
    cConexCom.Execute "ALTER TABLE CONFIGURACION  ADD COLUMN  cosven_Habe Text(10) "
 End If
 
 ProgressBar1.Value = ProgressBar1.Value + 1
 DoEvents
 
 If Not ExisteElem(1, cConexCom, "CONFIGURACION", "Alma_defa") Then
    cConexCom.Execute "ALTER TABLE CONFIGURACION  ADD COLUMN  Alma_defa Text(10) "
 End If
 
 ProgressBar1.Value = ProgressBar1.Value + 1
 DoEvents
 
 If Not ExisteElem(1, cConexCom, "CONFIGURACION", "Ladrillera") Then
    cConexCom.Execute "ALTER TABLE CONFIGURACION  ADD COLUMN  Ladrillera Text(1) " 'S=Si or N=No
 End If
   
 ProgressBar1.Value = ProgressBar1.Value + 1
 DoEvents
 
 If Not ExisteElem(1, cConexCom, "CONFIGURACION", "TIPO_ALMA") Then 'ALMACEN VENTAS O ALMACEN SUMINISTROS
    cConexCom.Execute "ALTER TABLE CONFIGURACION  ADD COLUMN  TIPO_ALMA Text(1) "
    cConexCom.Execute "UPDATE CONFIGURACION SET TIPO_ALMA='V'"
 End If
  
 LblMsg = "Verificando Tablas Relacionadas con el Maestro de Articulos  ..."
 ProgressBar1.Value = ProgressBar1.Value + 1
 DoEvents
 
'*****************************************************************
 If Not ExisteElem(0, cConexCom, "MAECOLOR") Then
       SQL = " Create Table MAECOLOR (COD_COLOR Text(20),DESCRI_COLOR Text(20), " & _
       " CONSTRAINT Clave PRIMARY KEY (COD_COLOR))"
       cConexCom.Execute SQL
 End If
 
 ProgressBar1.Value = ProgressBar1.Value + 1
 DoEvents
 
 If Not ExisteElem(0, cConexCom, "MAEMARCA") Then
       SQL = " Create Table MAEMARCA (COD_MARCA Text(20),DESCRI_MARCA Text(20), " & _
       " CONSTRAINT Clave PRIMARY KEY (COD_MARCA))"
       cConexCom.Execute SQL
 End If
 
 LblMsg = "Verificando Tabla de Ubicación  ..."
 ProgressBar1.Value = ProgressBar1.Value + 1
 DoEvents
 
 If Not ExisteElem(0, cConexCom, "TABUBICA") Then
       SQL = " Create Table  TABUBICA (COD_ALMA Text(2),COD_UBIC Text(20),DESCRI Text(45), " & _
       " CONSTRAINT Clave PRIMARY KEY (COD_ALMA,COD_UBIC))"
       cConexCom.Execute SQL
 End If
 
 LblMsg = "Verificando Tabla de CASILLEROS  ..."
 ProgressBar1.Value = ProgressBar1.Value + 1
 DoEvents
 
 If Not ExisteElem(0, cConexCom, "TABCASILLERO") Then
       SQL = " Create Table  TABCASILLERO (TCODALM Text(2),TCASILLERO Text(12),TCODART Text(20), " & _
       " CONSTRAINT Clave PRIMARY KEY (TCODALM,TCASILLERO,TCODART))"
       cConexCom.Execute SQL
 End If
 
 LblMsg = "Verificando Tabla de Temporal  ..."
 ProgressBar1.Value = ProgressBar1.Value + 1
 DoEvents
 
 If Not ExisteElem(0, cConexCom, "ART_LOTE") Then
    SQL = " Create Table ART_LOTE ( ALMA Text(3),ACODIGO Text(20),LOTE Text(20), CANTID DOUBLE )"
    cConexCom.Execute SQL
 End If

 ProgressBar1.Value = ProgressBar1.Value + 1
 LblMsg = "Verificando Tabla de Movimientos  ..."
 DoEvents
 
 If Not ExisteElem(1, cConexCom, "MOVALMDET", "DEORDFAB") Then
    cConexCom.Execute "ALTER TABLE MOVALMDET ADD COLUMN DEORDFAB Text(10) "
 End If
 
 If Not ExisteElem(1, cConexCom, "MOVALMDET", "DEQUIPO") Then
    cConexCom.Execute "ALTER TABLE MOVALMDET ADD COLUMN DEQUIPO Text(10) "
 End If
 
 LblMsg = "Verificando Tabla de Transacciones  ..."
 ProgressBar1.Value = ProgressBar1.Value + 1
 DoEvents
 
 If Not ExisteElem(1, cConexCom, "TABTRANSA", "TT_ORDFAB") Then
    cConexCom.Execute "ALTER TABLE TABTRANSA ADD COLUMN TT_ORDFAB Text(1) "
    cConexCom.Execute "Update TABTRANSA SET TT_ORDFAB='N'" 'INICIALIZA
 End If
 
 If Not ExisteElem(1, cConexCom, "TABTRANSA", "TT_EQUIP") Then
    cConexCom.Execute "ALTER TABLE TABTRANSA ADD COLUMN TT_EQUIP Text(1) "
    cConexCom.Execute "Update TABTRANSA SET TT_EQUIP='N'" 'INICIALIZA
 End If
 
' lblmsg = "Verificando Indices  ..."
' ProgressBar1.Value = ProgressBar1.Value + 1
' DoEvents
 
' If ExisteIndice(cRuta2, "InveFisiDet", "clave") Then
'    EliminaIndice cRuta2, "InveFisiDet", "clave"
'    CreaIndice cRuta2, "InveFisiDet", "PRIMARYKEY", True, "AUXNUMINVE", "AUXALMA", "AUXCODART"
' End If
 LblMsg = "Verificando Tablas para Centro de Costo..."
 ProgressBar1.Value = ProgressBar1.Value + 1
 DoEvents
    
 If Not ExisteElem(0, cConexCom, "MAQUINAS") Then
       SQL = " Create Table  MAQUINAS (MAQ_ALMA TEXT(2),COD_MAQ Text(10),DESC_MAQ Text(45),RESPON_MAQ Text(10),NHORA_DISP LONG,NHORA_TRAB LONG,ESTADO TEXT(5),CCOSTO_MAQ TEXT(10),FECH_CREA DATETIME,FECH_MODIF DATETIME, CODI_OPER TEXT(15)," & _
       " CONSTRAINT Clave PRIMARY KEY (MAQ_ALMA ,COD_MAQ))"
       cConexCom.Execute SQL
 End If
    
 ProgressBar1.Value = ProgressBar1.Value + 1
 DoEvents
    
 If Not ExisteElem(0, cConexCom, "ORDE_FAB") Then
       SQL = " Create Table  ORDE_FAB (ORD_ALMA TEXT (2),ORD_FABNUM Text(10),ORD_CODCLIE Text(15),ORD_CODART Text(20), ORD_CANT DOUBLE,FECH_INI DATETIME,FECH_FIN DATETIME,FECH_TRAN DATETIME,CODI_OPER TEXT(15)," & _
       " CONSTRAINT Clave PRIMARY KEY (ORD_ALMA,ORD_FABNUM))"
       cConexCom.Execute SQL
 End If
    
 LblMsg = "Verificando chequeos de longitud cero..."
 ProgressBar1.Value = ProgressBar1.Value + 1
 DoEvents
 
 If Not ExisteElem(1, cConexCom, "STKLOTE", "STOBSERVA") Then
    cConexCom.Execute "ALTER TABLE STKLOTE  ADD COLUMN  STOBSERVA Text(255) "
 End If
 
 ProgressBar1.Value = ProgressBar1.Value + 1
 DoEvents
 
'  If Not ExisteElem(1, cConexCom, "FAMILIA", "FAM_COMPRA") Then
'      Conexion.Execute "ALTER TABLE " & familia & " ADD COLUMN  " & FAM_COMPRA & " TEXT(20)" '
'  End If
 
' Dim Dtb As Database
' Dim Tdf As TableDef
' Dim Campo As Field
'
' Set Dtb = OpenDatabase(cRuta2)
' Set Tdf = Dtb.TableDefs("MAEART")
' Tdf.Fields("AMARCA").AllowZeroLength = True
' Tdf.Fields("ACOLOR").AllowZeroLength = True
'
' ProgressBar1.Value = ProgressBar1.Value + 1
' DoEvents
'
' If ExisteElem(1, cConexCom, "InveFisiCab", "AUXOBSER") Then
'    Set Dtb = OpenDatabase(cRuta2)
'    Set Tdf = Dtb.TableDefs("InveFisiCab")
'    Tdf.Fields("AUXOBSER").AllowZeroLength = True
' End If
'
' ProgressBar1.Value = ProgressBar1.Value + 1
' DoEvents
'
' Set Dtb = OpenDatabase(cRuta2)
' Set Tdf = Dtb.TableDefs("TABTRANSA")
' Tdf.Fields("TT_ORDFAB").AllowZeroLength = True
'
' ProgressBar1.Value = ProgressBar1.Value + 1
' DoEvents
'
' Set Dtb = OpenDatabase(cRuta2)
' Set Tdf = Dtb.TableDefs("TABTRANSA")
' Tdf.Fields("TT_EQUIP").AllowZeroLength = True
'
' ProgressBar1.Value = ProgressBar1.Value + 1
' DoEvents
'
' Set Dtb = OpenDatabase(cRuta2)
' Set Tdf = Dtb.TableDefs("MOVALMDET")
' Tdf.Fields("DEORDFAB").AllowZeroLength = True
'
' ProgressBar1.Value = ProgressBar1.Value + 1
' DoEvents
'
' Set Dtb = OpenDatabase(cRuta2)
' Set Tdf = Dtb.TableDefs("MOVALMDET")
' Tdf.Fields("DEQUIPO").AllowZeroLength = True
'
' ProgressBar1.Value = ProgressBar1.Value + 1
' DoEvents
'
' Set Dtb = OpenDatabase(cRuta2)
' Set Tdf = Dtb.TableDefs("InveFisiDet")
' Tdf.Fields("AUXFAMIL").AllowZeroLength = True
'
' ProgressBar1.Value = ProgressBar1.Value + 1
' DoEvents
'
' Set Dtb = OpenDatabase(cRuta2)
' Set Tdf = Dtb.TableDefs("STKLOTE")
' Tdf.Fields("STOBSERVA").AllowZeroLength = True
'
' ProgressBar1.Value = ProgressBar1.Value + 1
' DoEvents
'
' Set Dtb = OpenDatabase(cRuta2)
' Set Tdf = Dtb.TableDefs("NUM_DOCUMENTOS")
' Tdf.Fields("CTIMPRESORA").AllowZeroLength = True
'
' ProgressBar1.Value = ProgressBar1.Value + 1
' DoEvents
'
' Set Dtb = OpenDatabase(cRuta2)
' Set Tdf = Dtb.TableDefs("NUM_DOCUMENTOS")
' Tdf.Fields("CTPTO").AllowZeroLength = True
'
' ProgressBar1.Value = ProgressBar1.Value + 1
' DoEvents
'
' Set Dtb = OpenDatabase(cRuta2)
' Set Tdf = Dtb.TableDefs("Configuracion")
' Tdf.Fields("Alma_defa").AllowZeroLength = True
'
' ProgressBar1.Value = ProgressBar1.Value + 1
' DoEvents
'
' Set Dtb = OpenDatabase(cRuta2)
' Set Tdf = Dtb.TableDefs("NUM_DOCUMENTOS")
' Tdf.Fields("CTSERNUM").AllowZeroLength = True
'
' ProgressBar1.Value = ProgressBar1.Value + 1
' DoEvents
'
' Set Dtb = OpenDatabase(cRuta2)
' Set Tdf = Dtb.TableDefs("NUM_DOCUMENTOS")
' Tdf.Fields("CTCAMBIO").AllowZeroLength = True
'
' ProgressBar1.Value = ProgressBar1.Value + 1
' DoEvents
'
' Set Dtb = OpenDatabase(cRuta2)
' Set Tdf = Dtb.TableDefs("NUM_DOCUMENTOS")
' Tdf.Fields("CTSTOCK").AllowZeroLength = True
'
' ProgressBar1.Value = ProgressBar1.Value + 1
' DoEvents
'
' Set Dtb = OpenDatabase(cRuta2)
' Set Tdf = Dtb.TableDefs("NUM_DOCUMENTOS")
' Tdf.Fields("CTCONTROLADOR").AllowZeroLength = True
'
' ProgressBar1.Value = ProgressBar1.Value + 1
' DoEvents
'
' Set Dtb = OpenDatabase(cRuta2)
' Set Tdf = Dtb.TableDefs("NUM_DOCUMENTOS")
' Tdf.Fields("CTPUERTO").AllowZeroLength = True
'
' ProgressBar1.Value = ProgressBar1.Value + 1
' DoEvents
'
' Set Dtb = OpenDatabase(cRuta2)
' Set Tdf = Dtb.TableDefs("CONFIGURACION")
' Tdf.Fields("cosven_Habe").AllowZeroLength = True
' ProgressBar1.Value = ProgressBar1.Value + 1
' DoEvents
'
' Set Dtb = OpenDatabase(cRuta2)
' Set Tdf = Dtb.TableDefs("CONFIGURACION")
' Tdf.Fields("cosven_debe").AllowZeroLength = True
'
' ProgressBar1.Value = ProgressBar1.Value + 1
' DoEvents
'
'
' Set Dtb = OpenDatabase(cRuta2)
' Set Tdf = Dtb.TableDefs("CONFIGURACION")
' Tdf.Fields("conf_codigoIng").AllowZeroLength = True
'
' Set Dtb = OpenDatabase(cRuta2)
' Set Tdf = Dtb.TableDefs("CONFIGURACION")
' Tdf.Fields("cod_Iasa").AllowZeroLength = True
'
' 'Fernando: 05/09/2001
' '*****Ultima Actualizacion FFCC caso Foresta**********************
' '31/08/2001
' 'Se crea una tabla talla y los campos TALLA PA
' '********
' If Not ExisteElem(0, cConexCom, "TALLA") Then
'    cConexCom.Execute "CREATE TABLE TALLA(CODIGO TEXT(3),DESCRIP TEXT(50),OBSERVA TEXT(50))"
'    Call DEMORA(6)
'    Call ModiFieldDef(cRuta2, "TALLA", "DESCRIP", , , True, False, "")
'    Call DEMORA(6)
'    Call ModiFieldDef(cRuta2, "TALLA", "OBSERVA", , , True, False, "")
' End If
' If Not ExisteElem(1, cConexCom, "MAEART", "PA") Then
'    cConexCom.Execute "ALTER TABLE MAEART ADD COLUMN PA TEXT(30)"
'    Call DEMORA(6)
'    Call ModiFieldDef(cRuta2, "MAEART", "PA", , , True, False, "")
' End If
' If Not ExisteElem(1, cConexCom, "MAEART", "TALLA") Then
'    cConexCom.Execute "ALTER TABLE MAEART ADD COLUMN TALLA TEXT(3)"
'    Call DEMORA(6)
'    Call ModiFieldDef(cRuta2, "MAEART", "TALLA", , , True, False, "")
' End If
' If Not ExisteIndice(cRuta2, "TALLA", "PrimaryKey") Then
'    Call CreaIndice(cRuta2, "TALLA", "PrimaryKey", True, "CODIGO")
' End If
'
' '*****Fin de actualizacion de fernando
' Dtb.Close
' Set Dtb = Nothing
 Screen.MousePointer = 1
 Exit Sub
ERRAR:
       MsgBox Err.Description
       Resume Next
       Resume
End Sub

Private Sub Conex_ExecuteComplete(ByVal RecordsAffected As Long, ByVal pError As adodb.error, adStatus As adodb.EventStatusEnum, ByVal pCommand As adodb.Command, ByVal pRecordset As adodb.Recordset, ByVal pConnection As adodb.Connection)
  DoEvents
End Sub

Private Sub Conex_WillExecute(Source As String, CursorType As adodb.CursorTypeEnum, LockType As adodb.LockTypeEnum, Options As Long, adStatus As adodb.EventStatusEnum, ByVal pCommand As adodb.Command, ByVal pRecordset As adodb.Recordset, ByVal pConnection As adodb.Connection)
  DoEvents
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
  UpdateDatabases.Caption = "Verificación de Datos de la Empresa"
  CmdInit.Visible = True
  central Me
'' Usa el tercer marco de la secuencia como marco de inicio.
'  Picture1.Picture = PictureClip2.GraphicCell(2)
'  Y = 2
'' Centra la imagen más pequeña en la imagen más grande.
'  Picture1.Left = (Picture6.ScaleWidth - Picture1.Width) / 2
'  Picture1.Top = (Picture6.ScaleHeight - Picture1.Height) / 2
'  Picture6.Visible = False
End Sub

Private Sub Timer1_Timer()
  If toggle = 1 Then runtop
End Sub
Private Sub ADOConectar()
Dim cRt As String
cRt = App.Path & "\BdAuxCom.Mdb"
Set cConexAux = New adodb.Connection
cConexAux.CursorLocation = adUseClient
cConexAux.ConnectionString = "PROVIDER=Microsoft.Jet.OLEDB.3.51;Data Source=" & cRt & ";"
cConexAux.Open
End Sub
