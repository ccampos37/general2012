Attribute VB_Name = "ModPlan"
    Dim SEGURIDAD As ProcSistema        'DEFINICIóN DE LA VARIABLES DEL SISTEMA

Public VGL_SERVER As String ' VARIABLE GLOBAL QUE ME INDICA CON QUE SERVIDOR FUNCIONA EL SISTEMA
Public VGL_SERVCONTA As String ' VARIABLE QUE ME INDICA EL SERVIDOR DE CONTABILIDAD
Public VGL_SERVERREP As String ' VARIABLE GLOBAL QUE ME INDICA CON QUE SERVIDOR FUNCIONA EL SISTEMA
Public VGL_SERVCONTAREP As String ' VARIABLE QUE ME INDICA EL SERVIDOR DE CONTABILIDAD
Public VGL_BASEPRINCIPAL As String ' ME INDICA LA BASE DE DATOS STARPLAN
Public VGL_BASEAUXILIAR As String
Public VGL_BASE As String 'EL NOMBRE DE LA BASE DE DATOS DE LA EMPRESA CONQUE ESTA TRABAJANDO
Public VGL_COMPUTER As String ' VARIABLE QUE ME INDICA EL NOMBRE DE LA COMPUTADORA
Public VGL_INTEGRWNT As Boolean 'VARIABLE ME DICE SI UTILIZAR POR LA AUTENTIFICACION NT

Public USUARIOINI As String ' VARIABLE QUE ME INDICA EL USURAIO PUESTO EN EL ARCHIVO INI
Public UNIDADLOGICA As String
Public SQL As String

Public VGL_DATE As String
Public VGL_USUARIO As String
Public VGL_LOGON As String

'SECCIóN PARA LA DECLARACIóN DE VARIABLES
Public DBSYSTEM As ADODB.Connection       'CONECCIóN DE LA PLANILLA POR EMPRESA
Public DBAUXCOM As New ADODB.Connection   'BASE DE DATOS AUXILIAR
Public DBADMINPER As New ADODB.Connection 'BASE DE DATOS DE RECURSOS HUMANOS
Public DBSTARPLAN As New ADODB.Connection 'CONEXION PARA LA PLANTILLA
Public C As Integer

Public DBAUXSQL As New ADODB.Connection

Public IDENTITY As Boolean

Public VGUTIL(2) As Variant
Public Const VPEMPRESA = "BRISATRADE SAC"
Public VPTAREA As String        'VARIABLE DE COMANDOS Y TAREAS
Public VPCODTMP As String       'VARIABLE DE CóDIGOS TEMPORALES
Public VPFECHA As Date
Public VPTRASPRM As String      'VARIABLE DE TRASPASO DE PARáMETROS
Public VPNUMTMP As String       'NUMERO AUXILIAR
Public ARRSEXO(2) As String
Public ARRSITUACION(9) As String
Public ARRESTCIVIL(4) As String
Public CLMENU As New ClassMenu
Public CLASINFO As New ComputerInfo
Public PROCSIS As New ProcSistema
Public VAR_SHOW As Integer 'PARA REFRESCAR EL FORMULARIOS DE EVENTO O ESTUDIOS DESDE PERSONAL
Public VAR_MODO_EDIT As Boolean
Public VAR As Integer
Public VGPARAMREP() As Variant, VGPARAMFORM As Variant
Global Const CGCADVAL As String = ",'""?;.&$"

Public Enum FORMAPLAN
    FILEBOLETA = 0
    FILEPLANILLA = 1
    FILEPLANCAB = 2
End Enum

Type TYPETRANS
    RUTABASE As String
    TABLA As String
    ORDENADO As String
    FIELDK  As String
    EMPRESA As String
    ESCADENA As Boolean
End Type
Public TRAS As TYPETRANS

Public xTrab, XTRABNOMBRE

Type TYPEREGSELECT
    FECHACESEMAX As Date
    FECHAINIMAX As Date
    FECHAINI As Date
    SITUACIONES As String
    TIPOSTRABS As String
    USARFECHACESE As Boolean
End Type
Public REGSELECT As TYPEREGSELECT

Type TYPEREGSISTEM
    TABLAADEL As String     'TABLA ACTIVA DE ADELANTOS (ANUAL)
    TABLAPLAN As String     'TABLA DE PLANILLAS (ANUAL)
    TABLADETPLAN As String  'TABLA DE DETALLES DE PLANILLAS (ANUAL)
    TABLAQUINTA As String   'TABLA CORRESPONDIENTE A LA 5TA. CATEGORIA
    ANNO As Integer         'AñO ACTUAL EN CURSO
    PASSWORD As String      'PASSWORD DEL USUARIO ACTUAL
    USER As String          'USUARIO ACTUAL EN EL SISTEMA
    ALMACEN As String       'ALMACEN DE DATOS PASADOS
    EMPRESA As String       'NOMBRE DE EMPRESA ACTUAL DEL SISTEMA
    DIRECCION As String
    PATH As String          'DIRECTORIO DEL SISTEMA
    REPORTES As String
    PATHEMPRESA As String
    PATHFOTOS As String
    VERSION As String       'VERSIóN DEL SISTEMA
    INICIO As String        'FORMULARIO DE INICIO
    BIBLIO As String        'BIBLIOTECA DE RECURSOS
    COLPLANADEL As String   'NOMBRE DE LA COLUMNA DE PLANILLAS DE ADELANTOS
    RUC As String * 11      'NúMERO DE RUC DE LA EMPRESA
    ARCHIVOWENTPL As String
    LOGIN As String
    '----------Sección de Contabilidad------------------
    scTieneStConta As Boolean       'si tiene o no el sistema de contabilidad de enterprise
    scRutaBDWenco As String         'Ruta del archivo de configuración de contabilidad
    scRutaEmpresaWenco As String    'Ruta de la empresa en wenco
    scTipoAnexo As String           'Tipo de anexo para trabajadores
    scSubdi As String               'Tipo de Subdiario
    scCuenta As String              'Cuenta para armar la planilla
    scCompro As String              'Comprobante a Generar
    scCtaRedon As String
    scCreaTrab As Boolean           'Crea Trabajadores en contabilidad
    scNivelCta As String            'Ultimo Nivel de Cuenta en el plan de cuentas
    
    VALRRHH As Boolean
    ESADMINISTRADOR As Boolean
    BASESQL As String
    CODUSERTMP As String
    FORMATOFECHA As String
End Type
Public REGSISTEMA As TYPEREGSISTEM

Type REGWIN                 'REGISTRO DE CONFIGURACIóN DE VENTANAS
    NUEVO As Boolean        'NEW
    EDITAR As Boolean       'EDIT
    ELIMINAR As Boolean     'DEL
    IMPRIMIR As Boolean     'PRINT
    PRELIMINAR As Boolean   'LIST
    BUSCAR As Boolean       'SEARCH
    FILTRAR As Boolean      'FILTER
End Type

Public VGLFRM As Integer

'USADOS EN INGRESO DE MOVIMIENTOS
Type TYPEREGINGMOV
    CODNOMBOL As Long
    FECHAINI As Date
    FECHAFIN As Date
    NOMBRE As String * 50
    AREA As Boolean
    CADCONDI As String
End Type
Public REGINGMOV As TYPEREGINGMOV

Type TYPEREGINPUT
    CENTROCOSTO As String
    MESACTIVO As Date
    Codigo As Long
    TIPOPLANILLA As Byte
    NOMBRE As String
    BOL_TABLE As String  'TABLA DE BOLETAS EJ. BOL042000
    MOV_TABLE As String 'MOVIMIENTOS DE BOLETAS
    CADENA As String
    FECHAINI As Date
    FECHAFIN As Date
    ACCION As String
    EXISTENBOLS As Boolean
    REDONDEO As Boolean
End Type
Public REGINPUT As TYPEREGINPUT

Type TYPEREGPROC
    CONTINUAR As Boolean
    ADELANTOS As Boolean
    PRESTAMOS As Boolean
    PRESEDIT As Boolean
    PRESNO As Boolean
    PRESMAX As Boolean
    Quinta As Boolean
    INGRESOS As Boolean
    BLANCO As Boolean
    NUEVOSADEL As Boolean
End Type
Public REGPROC As TYPEREGPROC

Type TYPEREGPLAN
    MES As Date
    FECHA As Date
    TABLABOL As String
    TABLAMOV As String
    AUTOR As String
    DATABASE As String
End Type

Public AMESES(13) As String

Type DATOTRAB
    CodigoTrab As String
    DOCUMENTO As String
    TIPODOCUMENTO As String
    ApePat As String
    ApeMat As String
    NOMBRE As String
    FechaNac As Date
    ESTADOCIVIL As String
    UBIGEO As String
    DIRECCION As String
    TELEFONO As String
    Sexo  As Integer
    FECHAING As Date
    SITUACION As String
    CARGO As String
    BASICO As String
    NFICHA As String
    CSEGURO As String
    DEVENGUE As String
    SALUDVIDA As String
    ASIGFAM As String
    FECHACESE As String
    ESTADOINTERNO As String
    CODIGOALTERNO As String
    BANCO As String
    CTACTE As String
    CTABANCO As String
    BANCOCTS As String
    AFP As String
    CUSPP As String
    AREA As String
    CCosto As String
    TIPOTRAB As String
    DEPARTAMENTO As String
    CENTROAR As String
End Type
Public DATOTRABAJADOR As DATOTRAB
Global LlamaFrm As Integer
Global QQUIEBRE As Integer
Public MAXIMOQUIEBRE As Integer


Public Sub Main()
On Error GoTo ERR
    If App.PrevInstance Then
        MsgBox "Ya existe un módulo del sistema cargado actualmente en sus sistema", vbCritical
        End
    End If
'Datos a eliminar cuando este concluido la entrada al sistema
    With REGSISTEMA
        .TABLAADEL = "ADEL2000"
        .ALMACEN = "NONE"
        .ANNO = 2000
        .BIBLIO = "NONE"
        .TABLAQUINTA = "QC2000"
        .INICIO = "UnKnow"
        .PASSWORD = "*****"
        .TABLADETPLAN = "DTPL2000"
        .TABLAPLAN = "PLAN2000"
        .USER = "Camtex"
        .VERSION = "1.0.01"
        .COLPLANADEL = "ADELANTO"
        .REPORTES = "C:\Prueba\Reportes"
        .PATHEMPRESA = "C:\Prueba\planilla"
        .PATHFOTOS = "C:\Prueba\planilla\Fotos"
        .EMPRESA = "BRISATRADE SAC"
        .RUC = "12345678901"
    End With
    Dim xCad As String
'ARCHIVO DE CONFIGURACION
    xCad = App.PATH & "\MARFICE.ini"
    If UCase(Dir$(xCad, vbArchive)) <> "MARFICE.INI" Then
        MsgBox "No se ha encontrado el archivo de inicialización del sistema", vbInformation
        End
    End If
    REGSISTEMA.PATH = sGetIni(xCad, "REPORTES", "PLANILLAS", "?")
    If REGSISTEMA.PATH = "?" Then
        MsgBox "No se Especifico la Ruta del Archivo MARFICE.INI donde se encuentra la Base de Datos SQL Principal ", vbCritical
        End
    End If
    REGSISTEMA.REPORTES = REGSISTEMA.PATH
    
'CAPTURA EL SERVIDOR
    'SE CAPTURA EL NOMBRE DE LA COMPUTADORA

    If sGetIni(xCad, "PLANILLAS", "TOPIMP", "0") = 0 Then MDIPrincipal.Toolbar1.Buttons("OTRO").ButtonMenus("IMPOPROD").Visible = False
    
    VGL_COMPUTER = Trim(CLASINFO.ComputerName)
    VGL_SERVER = UCase(sGetIni(xCad, "PLANILLAS", "SERVIDOR", "?"))
    If VGL_SERVER = "?" Then
        MsgBox "Ud. no ha definido el NOMBRE del servidor de Base de Datos. Verifique el Archivo de Configuración."
        End
    End If
    
    VGL_SERVCONTA = UCase(sGetIni(xCad, "PLANILLAS", "SERVIDORCONTA", "?"))
    If VGL_SERVCONTA = "?" Then
        MsgBox "Ud. no ha definido el NOMBRE del servidor de Base de Datos para contabilidad. Verifique el Archivo de Configuración."
        End
    End If
    VGL_BASE = UCase(sGetIni(xCad, "PLANILLAS", "BDPRINCIPAL", "STARPLAN"))
    VGL_INTEGRWNT = UCase(sGetIni(xCad, "PLANILLAS", "INTEGRWNT", False))
    Dim TablaTmp As String
    
    '**X PrsGetIni
    TablaTmp = "tempdb"
    
    TablaTmp = UCase(sGetIni(xCad, "PLANILLAS", "TMP", "?"))
        
    If TablaTmp = "?" Then
        TablaTmp = "tempdb"
    End If
        UNIDADLOGICA = UCase(sGetIni(xCad, "PLANILLAS", "RUTASERVIDOR", "?"))
        If UNIDADLOGICA = "?" Then
          MsgBox "La unidad logica no es admitida. Verifique el Archivo de Configuración."
          End
        End If
           
        USUARIOINI = UCase(sGetIni(xCad, "PLANILLAS", "USUARIO", "?"))
        If UNIDADLOGICA = "?" Then
          MsgBox "el Usuario no es admitido. Verifique el Archivo de Configuración."
          End
        End If
        
        REGSISTEMA.FORMATOFECHA = UCase(sGetIni(xCad, "PLANILLAS", "FECHA", "MDY"))
        If REGSISTEMA.FORMATOFECHA <> "DMY" And REGSISTEMA.FORMATOFECHA <> "MDY" Then REGSISTEMA.FORMATOFECHA = "MDY"
        
    Dim X As Integer
'Esta seccion es para que cuando se trabaje en modo de Diseño ya no se vea la pantalla
'de presentacion y ademas para que no pida la clave de acceso al sistema
    If GetSetting(App.CompanyName, "Planillas", "Nando", "No") <> "Hola" Then
        'Present.Show 1
        Load MDIPrincipal
       ' Unload Present
    End If
'Abre la Base de Datos de Configuración



    With DBSTARPLAN
        .CursorLocation = adUseClient
        .CommandTimeout = 0
        .ConnectionTimeout = 0
        If VGL_INTEGRWNT Then
            .ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=" & VGL_BASE & ""
          Else
          '  .ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=" & VGL_BASE & ""
            .ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=SOPORTE;Password=SOPORTE;Initial Catalog=" & VGL_BASE & ";Data Source=" & VGL_SERVER
        End If
        .ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Password=;Initial Catalog=" & VGL_BASE & ";Data Source=" & VGL_SERVER
        .Open
    End With
        DBSTARPLAN.Execute "EXECUTE sp_dboption '" & VGL_BASE & "', 'select into/bulkcopy', 'TRUE'"
        Set DBADMINPER = DBSTARPLAN
        
        'VERIFICA LA FORMATO DE FECHA DEL SQL
        
        On Error GoTo ERRFECHA
        Dim FECHA As String
        FECHA = Date
        DBSTARPLAN.Execute "INSERT INTO ADELANTOS (FECHAING) VALUES (" & DateSQL(FECHA) & ")"
        
    frPanEmp.Show 1
'Abre la Base de Trabajo

    Set DBSYSTEM = New ADODB.Connection
    With DBSYSTEM
        .CursorLocation = adUseClient
        .CommandTimeout = 0
        .ConnectionTimeout = 0
        If VGL_INTEGRWNT Then
            .ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=" & REGSISTEMA.BASESQL & ""
          Else
         '  .ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=" & REGSISTEMA.BASESQL & ""
        '  .ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=SOPORTE;Password=SOPORTE;Initial Catalog=" & REGSISTEMA.BASESQL & ";Data Source=" & VGL_SERVER
           .ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Password=;Initial Catalog=" & REGSISTEMA.BASESQL & ";Data Source=" & VGL_SERVER
        End If
        .ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Password=;Initial Catalog=" & REGSISTEMA.BASESQL & ";Data Source=" & VGL_SERVER
        .Open
    End With
        DBSYSTEM.Execute "EXECUTE sp_dboption '" & REGSISTEMA.BASESQL & "', 'select into/bulkcopy', 'TRUE'"
    If UCase(Dir$(REGSISTEMA.PATHEMPRESA & "\ADMINPER.SQL")) <> "ADMINPER.SQL" Then REGSISTEMA.VALRRHH = False Else REGSISTEMA.VALRRHH = True
    Call AdjuntarServ(DBSYSTEM, VGL_SERVCONTA)
    
    'Llenando algunos parametros
    Dim RSEMPRESA As ADODB.Recordset
    Set RSEMPRESA = New ADODB.Recordset
    RSEMPRESA.Open "select * from empresa", DBSYSTEM, adOpenKeyset, adLockReadOnly
    With REGSISTEMA
        .DIRECCION = ESNULO(RSEMPRESA!DIRECCIÓN, "")
    End With
    
'Base de Datos Auxiliar donde se encuentran las Consultas para formar los Reportes
    With DBAUXCOM
        .CommandTimeout = 0
        .ConnectionTimeout = 0
        .CursorLocation = adUseClient
        If VGL_INTEGRWNT Then
           .ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=" & TablaTmp & ""
         Else
           .ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=" & TablaTmp & ""
          .ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Password=;Initial Catalog=" & TablaTmp & ";Data Source=" & VGL_SERVER
        End If
        .Open
    End With
    InitValores
    ActualizarSistema
    'ESTABLECER LOS PARAMETROS DE CONTABILIDAD EN EL REGISTRO DEL SISTEMA
    Call SETCFGCONTA
    VGL_SERVERREP = VGL_SERVER
    VGL_SERVCONTAREP = VGL_SERVCONTA
    VGL_SERVER = "[" & VGL_SERVER & "]"
    VGL_SERVCONTA = "[" & VGL_SERVCONTA & "]"
    
    MDIPrincipal.Show
    MDIPrincipal.Caption = "Planillas: " & REGSISTEMA.EMPRESA
    MDIPrincipal.BarraEstado.Panels(1).Text = REGSISTEMA.USER & " "
    CargaMesMax
Exit Sub
ERR:
    MsgBox "No se realizó la conexión con la Base de Datos principal. Verifique el archivo de configuración.|" & Chr(13) & Chr(10) & ERR.Description, vbCritical, "Error de Conexión"
    Resume Next
    End
ERRFECHA:
    MsgBox (Error)
     MsgBox "El Formato de la Fecha no es la Correcta verifique el Archivo de Inicio ", vbCritical, "Error"
    Resume Next
    End
End Sub
Public Sub AdjuntarServ(Cnx As ADODB.Connection, Servidor As String)
    On Error GoTo ErrAdjunt
        Cnx.Execute "Exec sp_addlinkedserver '" & Servidor & "'"
    Exit Sub
ErrAdjunt:
    Exit Sub
End Sub

Public Sub InitValores()
    'Array para los tipos de sexo
    ARRSEXO(0) = "Femenino"
    ARRSEXO(1) = "Masculino"
    'Array de condiciones laborales de los trabajadores
    ARRSITUACION(0) = "Activo o subsidiado(EPS/SP)"
    ARRSITUACION(1) = "Activo o subsidiado"
    ARRSITUACION(2) = "Baja (EPS/SP)"
    ARRSITUACION(3) = "Baja"
    ARRSITUACION(4) = "Licencia sin goce de haber (EPS/SP)"
    ARRSITUACION(5) = "Licencia sin goce de haber"
    ARRSITUACION(6) = "Baja con conceptos pendientes por liquidar (EPS/SP)"
    ARRSITUACION(7) = "Baja con conceptos pendientes por liquidar"
    ARRSITUACION(8) = "Activo"
    ARRESTCIVIL(0) = "Soltero"
    ARRESTCIVIL(1) = "Casado"
    ARRESTCIVIL(2) = "Conviviente"
    ARRESTCIVIL(3) = "Divorciado"
    'Array para los meses
    AMESES(0) = "Errado"
    AMESES(1) = "Enero"
    AMESES(2) = "Febrero"
    AMESES(3) = "Marzo"
    AMESES(4) = "Abril"
    AMESES(5) = "Mayo"
    AMESES(6) = "Junio"
    AMESES(7) = "Julio"
    AMESES(8) = "Agosto"
    AMESES(9) = "Setiembre"
    AMESES(10) = "Octubre"
    AMESES(11) = "Noviembre"
    AMESES(12) = "Diciembre"
End Sub

'Funcion que permite que le pases una cadena como
'por ejemplo: 720-7172730-45-0+ 45
'y te devuelve: 720717273045045 solo numeros
Public Function SoloNumeros(ByVal Dato As String) As String
    Dim xCad As String, X As Integer
    If Dato = "" Then
        SoloNumeros = ""
        Exit Function
    End If
    xCad = ""
    For X = 1 To Len(Dato)
        If IsNumeric(Mid(Dato, X, 1)) Then xCad = xCad & Mid(Dato, X, 1)
    Next
    SoloNumeros = xCad
End Function

Public Sub ActualizarSistema()
    'PROCEDIMIENTO QUE CONTROLA LAS ACTUALIZACIONES
    'DE LA BASE DE DATOS PARA LA VERSIÓN ACTUAL
    '-------------------------------------------------------------
    On Error Resume Next
    Dim X As Integer
    Dim ACTUALIZADO As Boolean
    ACTUALIZADO = False
    Call ACTUCOLPLA
    If Not ExisteTabla("CONFIADEL") Then
        DBSYSTEM.Execute "CREATE TABLE CONFIADEL (CODIGO VARCHAR(10), NOMBRE VARCHAR(100), TIPO INTEGER NULL DEFAULT 1)"
        ACTUALIZADO = True
    End If
    
    If Not ExisteCampo("GENE", "FORMULASVAC", DBSYSTEM) Then
        DBSYSTEM.Execute "ALTER TABLE FORMULASVAC ADD GENE BIT"
        DBSYSTEM.Execute "UPDATE FORMULASVAC SET GENE =0"
        ACTUALIZADO = True
    End If
        
    If Not ExisteCampo("GENE", "FORMULASCTS", DBSYSTEM) Then
        DBSYSTEM.Execute "ALTER TABLE FORMULASCTS ADD GENE BIT"
        DBSYSTEM.Execute "UPDATE FORMULASCTS SET GENE =0"
        ACTUALIZADO = True
    End If
    
    'AGREGANDO LA SUMAAFP EN AFP REMUNRACIONES
    If Trim(ESNULO(DevuelveValor("SELECT AFPREMU FROM EMPRESA ", DBSYSTEM), "")) = "" Or Trim(ESNULO(DevuelveValor("SELECT AFPREMU FROM EMPRESA ", DBSYSTEM), "")) = 0 Then
        DBSYSTEM.Execute "UPDATE EMPRESA SET AFPREMU='SUMAAFP'"
        ACTUALIZADO = True
    End If
    
    If Not ExisteCampo("GENE", "FORMULASGRATI", DBSYSTEM) Then
        DBSYSTEM.Execute "ALTER TABLE FORMULASGRATI ADD GENE BIT"
        DBSYSTEM.Execute "UPDATE FORMULASGRATI SET GENE =0"
        ACTUALIZADO = True
    End If
    If Not ExisteTabla("OFICIAL5TA") Then
        DBSYSTEM.Execute "CREATE TABLE OFICIAL5TA (PERIODO INT, CODTRAB VARCHAR(8), INUMBOL INT, REMUNERA  Numeric(20,2) , TRIBUTO  Numeric(20,2) , MES DATETIME, CODNOMBOL INT)"
        ACTUALIZADO = True
    End If
        
    If Not ExisteTabla("DETADEL") Then
        DBSYSTEM.Execute "CREATE TABLE DETADEL (CODTRAB VARCHAR(8), NOMBRE VARCHAR(100), MES DATETIME, CONCEPTO VARCHAR(50), MONTO  Numeric(20,2) , IE VARCHAR(1), TOTAL  Numeric(20,2) ,NOMBOL INT)"
        ACTUALIZADO = True
    End If
    If Not ExisteCampo("NOMBOL", "DETADEL", DBSYSTEM) Then
        DBSYSTEM.Execute "ALTER TABLE DETADEL ADD NOMBOL INT"
        ACTUALIZADO = True
    End If
    'CREAR UN CAMPO SECUENCIA PARA ALMACENAR LA SECUENCIA D CTA CTE. PROGRAMADA
    If Not ExisteCampo("SECUENCIA", "PAGOSCTA", DBSYSTEM) Then
        DBSYSTEM.Execute "ALTER TABLE PAGOSCTA ADD SECUENCIA INTEGER DEFAULT 0"
        DBSYSTEM.Execute "UPDATE PAGOSCTA SET SECUENCIA = 0"
        ACTUALIZADO = True
    End If
    'CREAR UN CAMPO SECUENCIA EN MOVICTA PARA ALMACENAR LA ULTIMA SECUENCIA UTILIZADA
    'EN EL PAGO EN CUENTAS CORRIENTES PROGRAMADAS
    If Not ExisteCampo("ULTSECU", "MOVICTA", DBSYSTEM) Then
        DBSYSTEM.Execute "ALTER TABLE MOVICTA ADD ULTSECU INTEGER DEFAULT 0"
        DBSYSTEM.Execute "UPDATE MOVICTA SET ULTSECU = 0"
        ACTUALIZADO = True
    End If
        
    If Not ExisteCampo("CODCONCEP", "DETADEL", DBSYSTEM) Then
        DBSYSTEM.Execute "ALTER TABLE DETADEL ADD CODCONCEP VARCHAR(15)"
        ACTUALIZADO = True
    End If
    
    If Not ExisteCampo("CRITERIO", "CONCEPTOS", DBSYSTEM) Then
        DBSYSTEM.Execute "ALTER TABLE CONCEPTOS ADD CRITERIO VARCHAR(200) NULL"
        ACTUALIZADO = True
    End If
    
    Dim I As Single
    For I = 1 To 12
      If Not ExisteCampo("MES" & Format(I, "00"), "CONFIG5TA", DBSYSTEM) Then
           DBSYSTEM.Execute "ALTER TABLE CONFIG5TA ADD MES" & Format(I, "00") & " VARCHAR(4) NULL DEFAULT '0' "
      End If
    Next I
    For I = 1 To 12
        If Not ExisteCampo("ACUMULA" & Format(I, "00"), "CONFIG5TA", DBSYSTEM) Then
            DBSYSTEM.Execute "ALTER TABLE CONFIG5TA ADD ACUMULA" & Format(I, "00") & " VARCHAR(4) NULL DEFAULT '0' "
        End If
    Next
    If Not ExisteTabla("HIST5TA") Then
        DBSYSTEM.Execute "CREATE TABLE HIST5TA (MES INT, ANNO VARCHAR(4), CODTRAB VARCHAR(8), NOMBRES VARCHAR(100), [TOTAL PERCIBIDO]  Numeric(20,2) , [PROYECTADO FIN AÑO]  Numeric(20,2) , [TOTAL RENTA PERCIBIR]  Numeric(20,2) , [RENTA AFECTA]  Numeric(20,2) , [IMPUESTO ANUAL]  Numeric(20,2) , [ACUMULADO]  Numeric(20,2) , SALDO  Numeric(20,2) , [MONTO RETENER]  Numeric(20,2) , [RENTENCION ANTERIOR]  Numeric(20,2) )"
        ACTUALIZADO = True
    End If
    If Not ExisteTabla("FRACGRATI") Then
        DBSYSTEM.Execute "CREATE TABLE FRACGRATI (MES INT, JULIO INT, DICIEMBRE INT)"
        ACTUALIZADO = True
    End If
    'SI EXISTE LA TABLA FORMULAS DE VACCACIONES PARA LAS PROVISIONES
    If Not ExisteTabla("FORVACPRO") Then
        DBSYSTEM.Execute "CREATE TABLE FORVACPRO (CODIGO INT IDENTITY (1, 1), NOMBRE VARCHAR(100), FORMULA VARCHAR(250), CRITERIO VARCHAR(240), TIPO INT, ACTIVO BIT NULL DEFAULT 0)"
        ACTUALIZADO = True
    End If
    'SI EXISTE LA TABLA FORMULAS DE GRATIFICACIONES PARA LAS PROVISIONES
    If Not ExisteTabla("FORGRAPRO") Then
        DBSYSTEM.Execute "CREATE TABLE FORGRAPRO (CODIGO INT IDENTITY (1, 1), NOMBRE VARCHAR(100), FORMULA VARCHAR(250), CRITERIO VARCHAR(240), TIPO INT, ACTIVO BIT NULL DEFAULT 0)"
        ACTUALIZADO = True
    End If
    'SI EXISTE LA TABLA FORMULAS DE CTS PARA LAS PROVISIONES
    If Not ExisteTabla("FORCTSPRO") Then
        DBSYSTEM.Execute "CREATE TABLE FORCTSPRO (CODIGO INT IDENTITY (1, 1), NOMBRE VARCHAR(100), FORMULA VARCHAR(250), CRITERIO VARCHAR(240), TIPO INT, ACTIVO BIT NULL DEFAULT 0)"
        ACTUALIZADO = True
    End If
    
    If Not ExisteCampo("USARCRONOGRAMA", "EMPRESA", DBSYSTEM) Then
        DBSYSTEM.Execute "ALTER TABLE EMPRESA ADD USARCRONOGRAMA BIT NULL DEFAULT 0"
        ACTUALIZADO = True
    End If
    
    If Not ExisteCampo("CLASEBOLETA", "EMPRESA", DBSYSTEM) Then
        DBSYSTEM.Execute "ALTER TABLE EMPRESA ADD CLASEBOLETA BIT NULL DEFAULT 0"
        ACTUALIZADO = True
    End If
    'SI EXISTE EL CAMPO PERMITE PARA LA TABLA CONCEPTOS
    If Not ExisteCampo("PERMITE", "CONCEPTOS", DBSYSTEM) Then
        DBSYSTEM.Execute "ALTER TABLE CONCEPTOS ADD PERMITE BIT NULL DEFAULT 0"
        DBSYSTEM.Execute "UPDATE CONCEPTOS SET PERMITE=0"
        ACTUALIZADO = True
    End If
    
    
    If Not ExisteCampo("ACUMULAQUINTA", "TRABAJADORES", DBSYSTEM) Then
        DBSYSTEM.Execute "ALTER TABLE TRABAJADORES ADD ACUMULAQUINTA  Numeric(20,2)  NOT NULL DEFAULT 0"
        ACTUALIZADO = True
    End If
    If Not ExisteCampo("AFECTOQUINTA", "TRABAJADORES", DBSYSTEM) Then
        DBSYSTEM.Execute "ALTER TABLE TRABAJADORES ADD AFECTOQUINTA BIT NOT NULL DEFAULT 0"
        ACTUALIZADO = True
    End If
    If Not ExisteCampo("XREDONDEO", "TRABAJADORES", DBSYSTEM) Then
        DBSYSTEM.Execute "ALTER TABLE TRABAJADORES ADD XREDONDEO  Numeric(20,2) "
        DBSYSTEM.Execute "UPDATE TRABAJADORES SET XREDONDEO=0"
        ACTUALIZADO = True
    End If
    If Not ExisteCampo("PROGRAMADO", "HISTOVAC", DBSYSTEM) Then
        'ACTUALIZACIÓN SOPORTE
        DBSYSTEM.Execute "ALTER TABLE HISTOVAC ADD NOMBOL INT"
        DBSYSTEM.Execute "ALTER TABLE HISTOVAC ADD FORMADESCANSO BIT NOT NULL DEFAULT 0"
        DBSYSTEM.Execute "ALTER TABLE HISTOVAC ADD COLUMN DIASCOMPENSADOS BIT NOT NULL DEFAULT 0"
        DBSYSTEM.Execute "ALTER TABLE HISTOVAC ADD COLUMN MODOCALCULO BIT NOT NULL DEFAULT 0"
        DBSYSTEM.Execute "ALTER TABLE HISTOVAC ADD COLUMN PROGRAMADO BIT NOT NULL DEFAULT 0"
        ACTUALIZADO = True
    End If
    If Not ExisteTabla("DETALLEVAC") Then
        'ACTUALIZACIÓN SOPORTE REPORTE BOLETA ESP. TECSUR
        'USO A NIVEL STANDARD - FUNCIONA CON TODOS LOS CLIENTES
        DBSYSTEM.Execute "CREATE TABLE DETALLEVAC (CODIGO INT, DESCRIPCION VARCHAR(30), IMPORTE  Numeric(20,2)  NOT NULL DEFAULT 0)"
        ACTUALIZADO = True
    End If
    'PAGOS DE PROGRAMACIÓN DE MOVIMIENTOS DE CUENTAS CORRIENTES
    If Not ExisteTabla("CTACTEPROG") Then
        DBSYSTEM.Execute "ALTER TABLE MOVICTA ADD PROGRAMADO BIT NOT NULL DEFAULT 0"
        DBSYSTEM.Execute "CREATE TABLE CTACTEPROG (CODMOV INT, CODTRAB VARCHAR(8), SECUENCIA BIT NOT NULL DEFAULT 0, FECHA DATETIME, IMPORTE  Numeric(20,2)  NOT NULL DEFAULT 0, CODNOMBOL INT)"
        ACTUALIZADO = True
    End If
    If Not ExisteCampo("MITAPER", "CTACTEPROG", DBSYSTEM) Then
        'ESTO SE MODIFICO CON EL SISTEMA TECSUR, APLICABLE A TODOS - JAAA, SOLUCIONES BABY - SOLO SOLUCIONES!!
        DBSYSTEM.Execute "ALTER TABLE CTACTEPROG ADD MITAPER CHAR(1)"
        ACTUALIZADO = True
    End If
    
    If Not ExisteCampo("IMPRESIONFIJA", "CONCEPTOS", DBSYSTEM) Then
        'ESTO SE MODIFICO CON EL SISTEMA TECSUR, APLICABLE A TODOS - JAAA, SOLUCIONES BABY - SOLO SOLUCIONES!!
        DBSYSTEM.Execute "ALTER TABLE CONCEPTOS ADD IMPRESIONFIJA BIT NOT NULL DEFAULT 0"
        ACTUALIZADO = True
    End If
    If Not ExisteTabla("SUBAREAS") Then
        DBSYSTEM.Execute "CREATE TABLE SUBAREAS (CODIGO VARCHAR(4), DESCRIPCION VARCHAR(50))"
        ACTUALIZADO = True
    End If
    'AQUI SE CREA UN CAMPO PARA LOS REPORTES DE CABECERA DE PLANILLAS
    If Not ExisteCampo("FILEPLANCAB", "EMPRESA", DBSYSTEM) Then
        DBSYSTEM.Execute "ALTER TABLE EMPRESA ADD FILEPLANCAB VARCHAR(50)"
        ACTUALIZADO = True
    End If
    If Not ExisteCampo("TIPOCTS", "CONCEPTOS", DBSYSTEM) Then
        'ESTO SE MODIFICO CON EL SISTEMA TECSUR, APLICABLE A TODOS
        DBSYSTEM.Execute "ALTER TABLE CONCEPTOS ADD TIPOCTS BIT NOT NULL DEFAULT 1"
        DBSYSTEM.Execute "ALTER TABLE CONCEPTOS ADD TIPOVAC BIT NOT NULL DEFAULT 1"
        DBSYSTEM.Execute "ALTER TABLE CONCEPTOS ADD TIPOGRA BIT NOT NULL DEFAULT 1"
        ACTUALIZADO = True
    End If
    If Not ExisteCampo("INDCTS", "CONCEPTOS", DBSYSTEM) Then
        'ESTO SE MODIFICO CON EL SISTEMA TECSUR, APLICABLE A TODOS
        DBSYSTEM.Execute "ALTER TABLE CONCEPTOS ADD INDCTS BIT NOT NULL DEFAULT 0"
        DBSYSTEM.Execute "ALTER TABLE CONCEPTOS ADD INDVAC BIT NOT NULL DEFAULT 0"
        DBSYSTEM.Execute "ALTER TABLE CONCEPTOS ADD INDGRA BIT NOT NULL DEFAULT 0"
        ACTUALIZADO = True
    End If
    If Not ExisteTabla("CTS") Then
        DBSYSTEM.Execute "CREATE TABLE CTS (CODIGO INT IDENTITY (1,1), NOMBRE VARCHAR(50),SOLES  Numeric(20,2)  NOT NULL DEFAULT 0, DOLARES  Numeric(20,2)  NOT NULL DEFAULT 0, CERRADO BIT NOT NULL DEFAULT 0, FECHAINI DATETIME, FECHAFIN DATETIME)"
        DBSYSTEM.Execute "CREATE TABLE PLANCTS (CODIGO INT NOT NULL DEFAULT 0, CODTRAB VARCHAR(8), NOMBRES VARCHAR(35), IMPORTECTS  Numeric(20,2)  NOT NULL DEFAULT 0, MESES BIT NOT NULL DEFAULT 0, DIAS BIT NOT NULL DEFAULT 0, FECHAING DATETIME, BANCO VARCHAR(4), CTABANCO VARCHAR(30), FECHADEPOSITO DATETIME, CUSTODIA BIT NOT NULL DEFAULT 0, PAGOCUSTODIO DATETIME)"
        DBSYSTEM.Execute "CREATE TABLE DETALLECTS (CODIGO INT, CODTRAB VARCHAR(8), CONCEPTO VARCHAR(35), IMPORTE  Numeric(20,2)  NOT NULL DEFAULT 0, INDTIPO BIT NOT NULL DEFAULT 0)"
        ACTUALIZADO = True
    End If
    If Not ExisteTabla("GRATIFICACION") Then
        DBSYSTEM.Execute "CREATE TABLE GRATIFICACION (CODIGO INT IDENTITY(1,1), NOMBRE VARCHAR(50), SOLES  Numeric(20,2)  NOT NULL DEFAULT 0, CERRADO BIT NOT NULL DEFAULT 0, FECHAINI DATETIME, FECHAFIN DATETIME, PERIODO INT)"
        DBSYSTEM.Execute "CREATE TABLE PLANGRATI (CODIGO INT , CODTRAB VARCHAR(8), NOMBRES VARCHAR(35), IMPORTEGRATI  Numeric(20,2)  NOT NULL DEFAULT 0, MESES BIT NOT NULL DEFAULT 0, DIAS BIT NOT NULL DEFAULT 0, FECHAING DATETIME)"
        DBSYSTEM.Execute "CREATE TABLE DETALLEGRATI (CODIGO INT, CODTRAB VARCHAR(8), CONCEPTO VARCHAR(35), IMPORTE  Numeric(20,2)  NOT NULL DEFAULT 0, INDTIPO BIT NOT NULL DEFAULT 0)"
        DBSYSTEM.Execute "CREATE TABLE FORMULASGRATI (CODIGO INT IDENTITY(1,1), NOMBRE VARCHAR(50), FORMULA VARCHAR(240), CRITERIO VARCHAR(240), TIPO BIT NOT NULL DEFAULT 0)"
        ACTUALIZADO = True
    End If
    If Not ExisteTabla("FORMULASCTS") Then
        DBSYSTEM.Execute "CREATE TABLE FORMULASCTS (CODIGO INT IDENTITY(1,1), NOMBRE VARCHAR(50), FORMULA VARCHAR(240), CRITERIO VARCHAR(240), TIPO BIT NOT NULL DEFAULT 0)"
        ACTUALIZADO = True
    End If
    If Not ExisteTabla("FORMULASVAC") Then
        DBSYSTEM.Execute "CREATE TABLE FORMULASVAC (CODIGO INT IDENTITY(1,1), NOMBRE VARCHAR(50), FORMULA VARCHAR(240), CRITERIO VARCHAR(240), TIPO BIT NOT NULL DEFAULT 0)"
        ACTUALIZADO = True
    End If
    If Not ExisteTabla("TRABXCOSTO") Then
        DBSYSTEM.Execute "CREATE TABLE TRABXCOSTO (CODTRAB VARCHAR(8), CODCCOSTO VARCHAR(10), BASICO  Numeric(20,2)  NOT NULL DEFAULT 0)"
        ACTUALIZADO = True
    End If
    If Not ExisteCampo("AFECTOPRO", "FORMULASVAC", DBSYSTEM) Then
        DBSYSTEM.Execute "ALTER TABLE FORMULASVAC ADD AFECTOPRO BIT NOT NULL DEFAULT 0"
        ACTUALIZADO = True
    End If
    If Not ExisteCampo("AFECTOPRO", "FORMULASGRATI", DBSYSTEM) Then
        DBSYSTEM.Execute "ALTER TABLE FORMULASGRATI ADD AFECTOPRO BIT NOT NULL DEFAULT 0"
        ACTUALIZADO = True
    End If
    If Not ExisteCampo("AFECTOPRO", "FORMULASCTS", DBSYSTEM) Then
        DBSYSTEM.Execute "ALTER TABLE FORMULASCTS ADD AFECTOPRO BIT NOT NULL DEFAULT 0"
        ACTUALIZADO = True
    End If
    If Not ExisteTabla("UTIL") Then
        DBSYSTEM.Execute "CREATE TABLE DETALLEUTIL ( CODIGO INT ,CODTRAB VARCHAR(8),CONCEPTO VARCHAR(35),IMPORTE  Numeric(20,2)  NOT NULL DEFAULT 0)"
        DBSYSTEM.Execute "CREATE TABLE FORMULASUTIL ( CODIGO INT IDENTITY(1,1), NOMBRE VARCHAR(240),FORMULA VARCHAR(240), CRITERIO VARCHAR(240))"
        DBSYSTEM.Execute "CREATE TABLE UTIL (CODIGO INT IDENTITY(1,1) ,NOMBRE VARCHAR(50), NOMBOL INT NOT NULL DEFAULT 0, PDIASSOLES  Numeric(20,2)  NOT NULL DEFAULT 0, PDIASDOLARES  Numeric(20,2)  NOT NULL DEFAULT 0, PREMSOLES  Numeric(20,2)  NOT NULL DEFAULT 0, PREMDOLARES  Numeric(20,2)  NOT NULL DEFAULT 0, CERRADO BIT NOT NULL DEFAULT 0, FECHAINI DATETIME, FECHAFIN DATETIME, FECHAACEP DATETIME, CALPER BIT NOT NULL DEFAULT 0, DIAEFECT BIT NOT NULL DEFAULT 0, PORPART  Numeric(20,2)  NOT NULL DEFAULT 0, UTILIDAD  Numeric(20,2)  NOT NULL DEFAULT 0, PARTDIST  Numeric(20,2)  NOT NULL DEFAULT 0, TOTPER  Numeric(20,2)  NOT NULL DEFAULT 0, IMPXPER  Numeric(20,2)  NOT NULL DEFAULT 0, TOTREM  Numeric(20,2)  NOT NULL DEFAULT 0, IMPXREM  Numeric(20,2)  NOT NULL DEFAULT 0)"
        DBSYSTEM.Execute "CREATE TABLE PLANUTIL(CODIGO INT, CODTRAB VARCHAR(8) ,NOMBRES VARCHAR(35), PARTPER  Numeric(20,2)  NOT NULL DEFAULT 0, PARTREM  Numeric(20,2)  NOT NULL DEFAULT 0, TOTPART  Numeric(20,2)  NOT NULL DEFAULT 0, DIAS INT NOT NULL DEFAULT 0, HORAS INT NOT NULL DEFAULT 0, TOTREM  Numeric(20,2)  NOT NULL DEFAULT 0, FECHAING DATETIME)"
        DBSYSTEM.Execute "CREATE INDEX CODTRAB ON PLANUTIL  (CODTRAB)"
        ACTUALIZADO = True
    End If
    If Not ExisteCampo("REDONDEO", "PLAN2000", DBSYSTEM) Then
        DBSYSTEM.Execute " ALTER TABLE PLAN2000 ADD REDONDEO  Numeric(20,2)  NOT NULL DEFAULT 0"
        ACTUALIZADO = True
    End If
    If Not ExisteTabla("DIASGOCE") Then
        DBSYSTEM.Execute "CREATE TABLE DIASGOCE (CODIGO INT, FECHAINI DATETIME, FECHAFIN DATETIME, DIAS BIT NOT NULL DEFAULT 0, DESCRIPCION VARCHAR(50))"
        ACTUALIZADO = True
    End If
    If Not ExisteCampo("M2", "LIQUIDACIONES", DBSYSTEM) Then DBSYSTEM.Execute "DROP TABLE LIQUIDACIONES"
    If Not ExisteTabla("LIQUIDACIONES") Then
        DBSYSTEM.Execute "CREATE TABLE LIQUIDACIONES (CODIGO INT IDENTITY(1,1), CODTRAB VARCHAR(8), FECHAING DATETIME, CARGO VARCHAR(35), FECHACESE DATETIME, BASECTS  Numeric(20,2)  NOT NULL DEFAULT 0, BASEVAC  Numeric(20,2) , BASEGRATI  Numeric(20,2)  NOT NULL DEFAULT 0, FECCTS DATETIME, FECVAC DATETIME, FECGRATI DATETIME, A1  Numeric(20,2)  NOT NULL DEFAULT 0, A2  Numeric(20,2)  NOT NULL DEFAULT 0, A3  Numeric(20,2)  NOT NULL DEFAULT 0, M1  Numeric(20,2)  NOT NULL DEFAULT 0, M2  Numeric(20,2)  NOT NULL DEFAULT 0, M3  Numeric(20,2)  NOT NULL DEFAULT 0, D1  Numeric(20,2)  NOT NULL DEFAULT 0, D2  Numeric(20,2)  NOT NULL DEFAULT 0, D3  Numeric(20,2)  NOT NULL DEFAULT 0, CODAFP VARCHAR(2) ,AFP1  Numeric(20,2)  NOT NULL DEFAULT 0, AFP2  Numeric(20,2)  NOT NULL DEFAULT 0, NETO  Numeric(20,2)  NOT NULL DEFAULT 0)"
        DBSYSTEM.Execute "CREATE TABLE DETALLELIQ (CODIGO INT, DESCRIPCION VARCHAR(35), IMPORTE  Numeric(20,2)  NOT NULL DEFAULT 0, TIPO BIT NOT NULL DEFAULT 0)"
        ACTUALIZADO = True
    End If
    If Not ExisteCampo("AFECTOPRO", "FORVACPRO", DBSYSTEM) Then
        DBSYSTEM.Execute "ALTER TABLE FORVACPRO ADD AFECTOPRO BIT NOT NULL DEFAULT 0"
        ACTUALIZADO = True
    End If
    If Not ExisteCampo("AFECTOPRO", "FORGRAPRO", DBSYSTEM) Then
        DBSYSTEM.Execute "ALTER TABLE FORGRAPRO ADD AFECTOPRO BIT NOT NULL DEFAULT 0"
        ACTUALIZADO = True
    End If
    If Not ExisteCampo("AFECTOPRO", "FORCTSPRO", DBSYSTEM) Then
        DBSYSTEM.Execute "ALTER TABLE FORCTSPRO ADD AFECTOPRO BIT NOT NULL DEFAULT 0"
        ACTUALIZADO = True
    End If
    
    'VERIFICO QUE EL CAMPO NOPDT NO ESTE EN NULOS
    DBSYSTEM.Execute "UPDATE TRABAJADORES SET NOPDT=0 WHERE (NOPDT NOT IN (1,0) OR NOPDT IS NULL)"
    'CREAR EL CAMPO NOCALCULO
    If Not ExisteCampo("NOCALCULO", "TRABAJADORES", DBSYSTEM) Then
        DBSYSTEM.Execute " ALTER TABLE TRABAJADORES ADD NOCALCULO BIT NOT NULL DEFAULT 0"
        ACTUALIZADO = True
    End If
    If Not ExisteCampo("NOPDT", "TRABAJADORES", DBSYSTEM) Then
        DBSYSTEM.Execute "ALTER TABLE TRABAJADORES ADD NOPDT BIT NOT NULL DEFAULT 0"
        DBSYSTEM.Execute "ALTER TABLE TRABAJADORES ADD OPCION01 BIT NOT NULL DEFAULT 0"
        DBSYSTEM.Execute "ALTER TABLE TRABAJADORES ADD OPCION02 BIT NULL DEFAULT 0"
        DBSYSTEM.Execute "ALTER TABLE TRABAJADORES ADD OPCIONA VARCHAR(15) NULL DEFAULT ''"
        DBSYSTEM.Execute "ALTER TABLE TRABAJADORES ADD OPCIONB VARCHAR(15) NULL DEFAULT ''"
        ACTUALIZADO = True
    End If
Dim Z As Boolean
Dim Y As Boolean
Z = False
    If Not ExisteCampo("TOTALEXTRA", "TRABAJADORES", DBSYSTEM) Then
        DBSYSTEM.Execute "ALTER TABLE TRABAJADORES ADD TOTALEXTRA  Numeric(20,2)  NOT NULL DEFAULT 0"
        Z = True
        ACTUALIZADO = True
    End If
Y = False
    If Not ExisteCampo("CERRADO", "MESESACT", DBSYSTEM) Then
        DBSYSTEM.Execute "ALTER TABLE MESESACT ADD CERRADO BIT NOT NULL DEFAULT 0"
        Y = True
        ACTUALIZADO = True
    End If
    If Not ExisteCampo("BKP_RUTA", "EMPRESA", DBSYSTEM) Then
        DBSYSTEM.Execute "ALTER TABLE EMPRESA ADD BKP_RUTA VARCHAR(250)"
        DBSYSTEM.Execute "ALTER TABLE EMPRESA ADD BKP_PERIODO BIT NOT NULL DEFAULT 0"
        DBSYSTEM.Execute "ALTER TABLE EMPRESA ADD BKP_DIASSEMANA BIT NOT NULL DEFAULT 0"
        DBSYSTEM.Execute "ALTER TABLE EMPRESA ADD BKP_DIASTRANS BIT NOT NULL DEFAULT 0"
        DBSYSTEM.Execute "ALTER TABLE EMPRESA ADD BKP_NUMBACKUPS BIT NOT NULL DEFAULT 0"
        ACTUALIZADO = True
    End If
    If Not ExisteCampo("COMENTARIO", "CONCEPTOS", DBSYSTEM) Then
        DBSYSTEM.Execute "ALTER TABLE CONCEPTOS ADD COMENTARIO VARCHAR(250)"
        DBSYSTEM.Execute "UPDATE CONCEPTOS SET COMENTARIO=''"
        ACTUALIZADO = True
    End If
    'ULTIMA ACTUALIZACION DE EMERRGENCIA: POR LIQUIDACIONES
    If Not ExisteCampo("CRONOVAC", "LIQUIDACIONES", DBSYSTEM) Then
        DBSYSTEM.Execute "ALTER TABLE LIQUIDACIONES ADD CRONOVAC INT"
        DBSYSTEM.Execute "ALTER TABLE LIQUIDACIONES ADD CRONOGRAT INT"
        ACTUALIZADO = True
    End If
    
    If Z Then DBSYSTEM.Execute "UPDATE TRABAJADORES SET TOTALEXTRA=0"
    
    If Not ExisteCampo("PROGRAMADO", "MOVICTA", DBSYSTEM) Then
        DBSYSTEM.Execute "ALTER TABLE MOVICTA ADD PROGRAMADO BIT NOT NULL DEFAULT 0"
        ACTUALIZADO = True
    End If
     If Not ExisteCampo("ADELVAC", "EMPRESA", DBSYSTEM) Then
        DBSYSTEM.Execute "ALTER TABLE EMPRESA ADD ADELVAC VARCHAR(16) NULL DEFAULT '' "
        ACTUALIZADO = True
    End If
    
    
    If Not ExisteCampo("XREDONDEO", "TRABBOLETEAR", DBSYSTEM) Then
        DBSYSTEM.Execute "DROP VIEW TRABBOLETEAR"
        DBSYSTEM.Execute "CREATE VIEW TRABBOLETEAR AS SELECT TRABAJADORES.CODTRAB, [TRABAJADORES].[APEPAT] + ' ' + LTrim([TRABAJADORES].[APEMAT]) + ' ' + LTrim([TRABAJADORES].[NOMBRE]) AS NOMBRES, TRABAJADORES.CCOSTO, TRABAJADORES.BASICO, AFPS.NOMBRE AS NOMAFP, AFPS.APOROBLI, AFPS.SEGURO, AFPS.TOPESEGURO, AFPS.COMISIONRA, TRABAJADORES.MESDEVENGUE, TRABAJADORES.ASIGFAM, CENTROSAR.NOMBRE AS NOMSCTR, CENTROSAR.TASA, TRABAJADORES.DEPARTAMENTO, TRABAJADORES.FONDOPENS, TRABAJADORES.AREA, TRABAJADORES.UBIGEO, TRABAJADORES.SEXO, TRABAJADORES.TIPOTRAB, TRABAJADORES.FECHAING, TRABAJADORES.SITUACIÓN, TRABAJADORES.CARGO, TRABAJADORES.BANCO, TRABAJADORES.ESSALUDVIDA, TRABAJADORES.RUCEPS, TRABAJADORES.NOPDT, TRABAJADORES.OPCION01, TRABAJADORES.OPCION02, TRABAJADORES.OPCIONA, TRABAJADORES.OPCIONB,TRABAJADORES.NOCALCULO, TRABAJADORES.XREDONDEO " & _
                         "FROM AFPS INNER JOIN (TRABAJADORES INNER JOIN CENTROSAR ON TRABAJADORES.CODSCTR = CENTROSAR.CODCAR) ON AFPS.CODAFP = TRABAJADORES.FONDOPENS "
    End If
    If Not ExisteCampo("AFECTOQUINTA", "TRABBOLETEAR", DBSYSTEM) Then
        DBSYSTEM.Execute "DROP VIEW TRABBOLETEAR"
        DBSYSTEM.Execute "CREATE VIEW TRABBOLETEAR AS SELECT TRABAJADORES.CODTRAB, [TRABAJADORES].[APEPAT] + ' ' + LTrim([TRABAJADORES].[APEMAT]) + ' ' + LTrim([TRABAJADORES].[NOMBRE]) AS NOMBRES, TRABAJADORES.CCOSTO, TRABAJADORES.BASICO, AFPS.NOMBRE AS NOMAFP, AFPS.APOROBLI, AFPS.SEGURO, AFPS.TOPESEGURO, AFPS.COMISIONRA, TRABAJADORES.MESDEVENGUE, TRABAJADORES.ASIGFAM, CENTROSAR.NOMBRE AS NOMSCTR, CENTROSAR.TASA, TRABAJADORES.DEPARTAMENTO, TRABAJADORES.FONDOPENS, TRABAJADORES.AREA, TRABAJADORES.UBIGEO, TRABAJADORES.SEXO, TRABAJADORES.TIPOTRAB, TRABAJADORES.FECHAING, TRABAJADORES.SITUACIÓN, TRABAJADORES.CARGO, TRABAJADORES.BANCO, TRABAJADORES.ESSALUDVIDA, TRABAJADORES.RUCEPS, TRABAJADORES.NOPDT, TRABAJADORES.OPCION01, TRABAJADORES.OPCION02, TRABAJADORES.OPCIONA, TRABAJADORES.OPCIONB,TRABAJADORES.NOCALCULO, TRABAJADORES.XREDONDEO, TRABAJADORES.AFECTOQUINTA " & _
                         "FROM AFPS INNER JOIN (TRABAJADORES INNER JOIN CENTROSAR ON TRABAJADORES.CODSCTR = CENTROSAR.CODCAR) ON AFPS.CODAFP = TRABAJADORES.FONDOPENS "
    End If


    If Not ExisteCampo("FIJO", "TABLREP", DBSTARPLAN) Then
        DBSTARPLAN.Execute "ALTER TABLE TABLREP ADD FIJO BIT NOT NULL DEFAULT 0"
        ACTUALIZADO = True
    End If
    
    'CREANDO LAS TABLAS DE CONTABILIDAD PARA EL PASE A CONTABILIDAD Y LA TABLA DE CONFIGURACIONES
    Call CREARTABLACONTABILIDAD
    Call CREACFGASIENTOS
    'NO SQL
    If Not ExisteTabla("CTACONCEPTO") Then
        DBSYSTEM.Execute "CREATE TABLE CTACONCEPTO(COD int IDENTITY (1, 1) NOT NULL,SEC INT,TIPOCTA VARCHAR(1),CUENTA VARCHAR(25),TIPASI INT,TIPASINOM VARCHAR(25),CONCEPT VARCHAR(8))"
        ACTUALIZADO = True
    End If
    'ACTUALIZAR EL TAMAÑO DEL CAMPO
    If TamaCampo("NOMBRES", "PLANCTS", DBSYSTEM) <= 36 Then
        DBSYSTEM.Execute "ALTER TABLE PLANCTS ALTER COLUMN NOMBRES VARCHAR(100)"
        ACTUALIZADO = True
    End If
    'ACTUALIZAR LOS CONCEPTOS DE CUENTAS COMO ADELANTOS(XXADELX), PAGOSCTA IGRESO(XXPAGCXI)
    'PAGOSCTA EGRESO(XXPAGCXE)
    If Not DevuelveValor("SELECT CODIGO FROM CONCEPTOS WHERE CODIGO='XXADELX'", DBSYSTEM) = "XXADELX" Then
        DBSYSTEM.Execute "" & _
        "INSERT INTO CONCEPTOS(CODIGO,NOMBRE,TIPO,FILA,ESESCRITO,TIPOINFO,TIPOREMU,FLAG) " & _
        "SELECT 'XXADELX','Para Cuenta de Adelantos',2,1,1,5,2,1 " ' UNION ALL " & _
'        "SELECT 'XXPAGCXI','Cuenta Corriente Ingreso',1,1,1,5,2,1 UNION ALL " & _
'        "SELECT 'XXPAGCXE','Cuenta Corriente Egresos',2,2,1,5,2,1 "
        ACTUALIZADO = True
    End If
    If TamaCampo("COMP", "CFGASIENTOS", DBSYSTEM) = 4 Then
        DBSYSTEM.Execute "ALTER TABLE CFGASIENTOS ALTER COLUMN COMP VARCHAR(25) "
        ACTUALIZADO = True
    End If
    If TamaCampo("RUC", "CCOSTOS", DBSYSTEM) = 8 Then
        DBSYSTEM.Execute "ALTER TABLE CCOSTOS ALTER COLUMN RUC VARCHAR(10) "
        ACTUALIZADO = True
    End If
    If TamaCampo("NOMBRES", "PLANGRATI", DBSYSTEM) = 35 Then
        DBSYSTEM.Execute "ALTER TABLE PLANGRATI ALTER COLUMN NOMBRES VARCHAR(100) "
        ACTUALIZADO = True
    End If
    If TamaCampo("CONCEPTO", "DETALLEGRATI", DBSYSTEM) = 35 Then
        DBSYSTEM.Execute "ALTER TABLE DETALLEGRATI ALTER COLUMN CONCEPTO VARCHAR(100) "
        ACTUALIZADO = True
    End If
    If Not ExisteTabla("MOTIVORENUNCIA") Then
        DBSYSTEM.Execute "CREATE TABLE MOTIVORENUNCIA(CODTRAB VARCHAR(8),MOTIVO VARCHAR(250)) "
        ACTUALIZADO = True
    End If
    
    If Not ExisteCampo("ULTMES", "NOMBOL", DBSYSTEM) Then
        DBSYSTEM.Execute "ALTER TABLE NOMBOL ADD ULTMES BIT NOT NULL DEFAULT 0"
        ACTUALIZADO = True
        DBSYSTEM.Execute "UPDATE NOMBOL SET ULTMES=0"
    End If
    'Actualizar Registros de Derechos Habientes su codigo
    Dim RsFamiliar As ADODB.Recordset
    Set RsFamiliar = New ADODB.Recordset
    Dim CONTDER As Long
    CONTDER = 1
    RsFamiliar.Open "FAMILIAR", DBSYSTEM, adOpenKeyset, adLockOptimistic
    If RsFamiliar.RecordCount > 0 Then
       If IsNull(RsFamiliar!CODDER) Then
           ACTUALIZADO = True
           Do While Not RsFamiliar.EOF
              RsFamiliar!CODDER = CONTDER
              RsFamiliar.Update
              RsFamiliar.MoveNext
              CONTDER = CONTDER + 1
           Loop
       End If
    End If
    'CREANDO UN REGISTRO PERMANENTE PARA COLOCAR EL TIPO DE CUENTA DE CADA TRABAJADOR
    If DevuelveValor("SELECT CODDATA FROM DATATRAB WHERE CODDATA='TIPCTAX'", DBSYSTEM) = "" Then
        ACTUALIZADO = True
        DBSYSTEM.Execute "INSERT INTO DATATRAB VALUES('TIPCTAX','Tipo de Cuenta del trabajador','T')"
    End If
    'CREANDO EL CAMPO TIPCTAX AHORA EN LA TABLA TRABAJADORES
    If Not ExisteCampo("TIPCTAX", "TRABAJADORES", DBSYSTEM) Then
        DBSYSTEM.Execute "ALTER TABLE TRABAJADORES ADD TIPCTAX VARCHAR(1) "
        ACTUALIZADO = True
        DBSYSTEM.Execute "UPDATE TRABAJADORES SET TIPCTAX=' '"
    End If
    'CREANDO UN REGISTRO PERMANENTE PARA COLOCAR LA CUENTA DESTINO
    If DevuelveValor("SELECT CODDATA FROM DATATRAB WHERE CODDATA='XXCTADES'", DBSYSTEM) = "" Then
        ACTUALIZADO = True
        DBSYSTEM.Execute "INSERT INTO DATATRAB VALUES('XXCTADES','Cuenta destino del trabajador','T')"
    End If
    'CREANDO EL CAMPO XXCTADES AHORA EN LA TABLA TRABAJADORES
    If Not ExisteCampo("XXCTADES", "TRABAJADORES", DBSYSTEM) Then
        DBSYSTEM.Execute "ALTER TABLE TRABAJADORES ADD XXCTADES VARCHAR(50) "
        ACTUALIZADO = True
        DBSYSTEM.Execute "UPDATE TRABAJADORES SET XXCTADES=''"
    End If
'*****************POR BASILIO
    
    If Not ExisteTabla("CONTQUIDET") Then
            DBSYSTEM.Execute "SELECT * INTO CONTQUIDET FROM CONTDET"
            DBSYSTEM.Execute "DELETE  FROM CONTQUIDET"
            DBSYSTEM.Execute "SELECT * INTO CONTQUICAB FROM CONTCAB"
            DBSYSTEM.Execute "DELETE  FROM CONTQUICAB"
    End If
    If Not ExisteTabla("CTACONCEPTOQUIN") Then
            DBSYSTEM.Execute "select * INTO CTACONCEPTOQUIN FROM CTACONCEPTO"
            DBSYSTEM.Execute "DELETE FROM CTACONCEPTOQUIN"
    End If
    If ExisteCampo("NETADEL", "CFGASIENTOS", DBSYSTEM) = False Then
        DBSYSTEM.Execute "alter table  CFGASIENTOS add  NETADEL VARCHAR(25),MONADEL VARCHAR(25),REDADEL VARCHAR(25)"
    End If
    
    Set CnAux = Nothing
    If ACTUALIZADO Then
        MsgBox "EL SISTEMA DE PLANILLAS HA PREPARADO LA BASE DE DATOS PARA SOPORTAR LA VERSIÓN ACTUAL", vbInformation
    End If
    
    If ActualizadoReporte Then
        MsgBox "EL SISTEMA DE PLANILLAS HA PREPARADO LAS CONSULTAS NECESARIAS PARA SOPORTAR LA VERSIÓN ACTUAL DEL SOFTWARE", vbInformation
    End If
    
End Sub
Public Function TamaCampo(CAMPO As String, TABLA As String, Cnx As ADODB.Connection) As Long
On Error GoTo ERRTAMA
    Dim RSAUX As New ADODB.Recordset
    RSAUX.Open "SELECT TOP 1 " & CAMPO & " FROM " & TABLA, Cnx
    TamaCampo = RSAUX(CAMPO).DefinedSize
    Exit Function
ERRTAMA:
    TamaCampo = 0
End Function
Public Function DevNomRep(EMPRESA As String, USUARIO As String, Formato As FORMAPLAN, Optional ByRef TipFmt) As String
    Dim RSAUX As New ADODB.Recordset
    DevNomRep = ""
    RSAUX.Open "SELECT * FROM TABLREP WHERE CODEMPRESA='" & REGSISTEMA.RUC & "' AND CODUSU='" & REGSISTEMA.USER & "' AND ACTIVO=-1", DBSTARPLAN
    If RSAUX.RecordCount = 0 Then Exit Function
    Select Case Formato
        Case 0: DevNomRep = RSAUX!FILEBOLETA
        Case 1: DevNomRep = RSAUX!FILEPLANILLA
        Case 2: DevNomRep = RSAUX!FILEPLANCAB
    End Select
    TipFmt = RSAUX!TipFmt
    DevNomRep = Trim(DevNomRep)
End Function

Public Sub VerificarSistema()
    If Not SEGURIDAD.PrVerifSeg Then
        MsgBox "No se ha registrado el sistema." & Chr(13) & Chr(10) & "Comuniquese con Enterprise Solutions S.A.", vbCritical
        End 'COMPROBACIÓN DEL SISTEMA (Copia)
    End If
End Sub
Public Sub CargaMesMax()
    On Error Resume Next
    Dim xFecha As Date
    If Not IsNull(DevuelveValor("SELECT MAX(MESACTIVO) FROM MESESACT", DBSYSTEM)) Then
        xFecha = DevuelveValor("SELECT MAX(MESACTIVO) FROM MESESACT", DBSYSTEM)
        BarraEstado.Panels("PERIODO").Text = Month(xFecha) & "/" & Year(xFecha)
    Else
        BarraEstado.Panels("PERIODO").Text = "Sin meses activos"
    End If
End Sub

Public Sub CambiaPanelBD(Estado As Boolean)
    If Estado Then
        Screen.MousePointer = 11
        MDIPrincipal.BarraEstado.Panels("BaseDatos").Text = "Espere ... en proceso "
        MDIPrincipal.BarraEstado.Panels("BaseDatos").Picture = MDIPrincipal.ImageList1.ListImages(25).ExtractIcon
    Else
        Screen.MousePointer = 1
        MDIPrincipal.BarraEstado.Panels("BaseDatos").Text = "Base de Datos Activa "
        MDIPrincipal.BarraEstado.Panels("BaseDatos").Picture = MDIPrincipal.ImageList1.ListImages(27).ExtractIcon
    End If
End Sub
Private Sub ACTUCOLPLA()
Dim VALOR As String
    VALOR = DevuelveValor("SELECT TOP 1 * FROM COLUMPL WHERE CODIGO='TOTING'", DBSYSTEM)
    If VALOR = "" Then DBSYSTEM.Execute "INSERT INTO COLUMPL  VALUES('TOTING','TOTAL INGRESOS','',4,2)"
        
    VALOR = DevuelveValor("SELECT TOP 1 * FROM COLUMPL WHERE CODIGO='TOTEGR'", DBSYSTEM)
    If VALOR = "" Then DBSYSTEM.Execute "INSERT INTO COLUMPL  VALUES('TOTEGR','TOTAL EGRESOS','',4,3)"
    
    VALOR = DevuelveValor("SELECT TOP 1 * FROM COLUMPL WHERE CODIGO='NETO'", DBSYSTEM)
    If VALOR = "" Then
        VALOR = DevuelveValor("SELECT TOP 1 * FROM COLUMPL WHERE CODIGO='NETOPAGO'", DBSYSTEM)
        If VALOR = "" Then DBSYSTEM.Execute "INSERT INTO COLUMPL  VALUES('NETO','NETO','',4,2)"
    End If
End Sub
Public Sub ACTSALDO(CODMOV As String, Optional MON As Integer)
'PROCEDIMIENTO QUE ACTULIZA LOS SALDOS DE CUENTA CORRIENTE UNA VEZ MODIFICADO
'O INGRESADO
'CREADO POR FERNANDO COSSIO
Dim X As Integer
Dim RSAUX As New ADODB.Recordset
Dim SALDO As Double
    If Not IsNumeric(ESNULO(CODMOV, 0)) Then Exit Sub
    RSAUX.Open "SELECT SUM(MONTO) AS TOTAL,SUM(DOLAR) AS SDOLAR FROM PAGOSCTA WHERE CODMOV=" & ESNULO(CODMOV, 0), DBSYSTEM
    If MON = 1 Then
        SALDO = ESNULO(RSAUX("SDOLAR"), 0)
      Else
        SALDO = ESNULO(RSAUX("TOTAL"), 0)
    End If
    DBSYSTEM.Execute "UPDATE MOVICTA SET SALDO=CAPITAL-" & SALDO & " WHERE CODMOV=" & ESNULO(CODMOV, 0), X
End Sub
Public Function Valc(CAD As Variant) As Double
'Creada por Fernando Cossio
'Funcion que a una cadena de numero ("1,514.89") devolvera el valor real en numeros
'no como la funcion tradicional Val que me devolveria 1 por que no considera a la coma parte
'del numero
On Error GoTo Errornum
    Valc = Val(Format(CAD, "#.######################################"))
    Exit Function
Errornum:
    Valc = 0
End Function
Private Sub CREARTABLACONTABILIDAD()
    'NO SQL
    'CREAR TABLAS PARA EL PASE CONTABLE
    Dim SqlCad As String
    'TABLA CABECERA
    SqlCad = "CREATE TABLE CONTCAB(NUM  int IDENTITY (1, 1) NOT NULL," & _
             "SUBDIAR_CODIGO  VARCHAR(4)," & _
             "CMOV_C_COMPR    VARCHAR(4)," & _
             "CMOV_FECHA  datetime, " & _
             "CMOV_GLOSA  VARCHAR(30)," & _
             "CMOV_MONED  VARCHAR(2)," & _
             "CMOV_CONVE  VARCHAR(3)," & _
             "CMOV_CAMES  FLOAT," & _
             "CMOV_FECCA  datetime, " & _
             "CMOV_TIPCA  FLOAT," & _
             "CMOV_DEBE   FLOAT, " & _
             "CMOV_HABER  FLOAT," & _
             "CMOV_DEBUS  FLOAT," & _
             "CMOV_HABUS  FLOAT," & _
             "CMOV_AUTOM  bit NOT NULL DEFAULT 0," & _
             "CMOV_COSTO  bit NOT NULL DEFAULT 0," & _
             "CMOV_CHEQU  bit NOT NULL DEFAULT 0," & _
             "CMOV_L_COMPR    bit NOT NULL DEFAULT 0," & _
             "CMOV_VENTA  bit NOT NULL DEFAULT 0,CRONO INT)"
    If Not ExisteTabla("CONTCAB") Then DBSYSTEM.Execute SqlCad
    'TABLA DETALLE
    SqlCad = "CREATE TABLE CONTDET(NUM INT," & _
             "SUBDIAR_CODIGO  VARCHAR(4), " & _
             "DMOV_C_COMPR    VARCHAR(4), " & _
             "DMOV_SECUE  VARCHAR(4), " & _
             "DMOV_FECHA  datetime," & _
             "DMOV_CUENT  VARCHAR(18)," & _
             "DMOV_ANEXO  VARCHAR(13)," & _
             "DMOV_DOCUM  VARCHAR(23)," & _
             "DMOV_FECDC  datetime," & _
             "DMOV_CENCO  VARCHAR(6)," & _
             "DMOV_DEBE   FLOAT, " & _
             "DMOV_HABER  FLOAT, " & _
             "DMOV_DEBUS  FLOAT, " & _
             "DMOV_HABUS  FLOAT," & _
             "DMOV_GLOSA  VARCHAR(30)," & _
             "DMOV_CHEQU  bit NOT NULL DEFAULT 0," & _
             "DMOV_AUTOM  bit NOT NULL DEFAULT 0," & _
             "DMOV_COSTO  bit NOT NULL DEFAULT 0," & _
             "DMOV_L_COMPR    bit NOT NULL DEFAULT 0," & _
             "DMOV_VENTA  bit NOT NULL DEFAULT 0," & _
             "DMOV_TRANS  bit NOT NULL DEFAULT 0," & _
             "DMOV_L_DESTI    bit NOT NULL DEFAULT 0," & _
             "DMOV_C_DESTI    VARCHAR(18),CRONO INT)"
     If Not ExisteTabla("CONTDET") Then DBSYSTEM.Execute SqlCad
End Sub
Private Sub CREACFGASIENTOS()
    'SE CREA LA TABLA DONDE SE GUARDA LA CONFIGURACION DE ASIENTOS DE PLANILLA PARA CONTABILIDAD
    Dim RS As New ADODB.Recordset
    Dim SqlCad As String
    SqlCad = "CREATE TABLE CFGASIENTOS(CHKCONTA BIT NOT NULL DEFAULT 0,RUTCONTA VARCHAR(150)," & _
           "CHKCREATRAB BIT NOT NULL DEFAULT 0,CODEMP VARCHAR(3),NOMEMP VARCHAR(50), TIPANEX VARCHAR(2)," & _
           "SUBDI VARCHAR(2),COMP VARCHAR(4),CUENTA VARCHAR(25))"
    If Not ExisteTabla("CFGASIENTOS") Then DBSYSTEM.Execute SqlCad
    RS.Open "CFGASIENTOS", DBSYSTEM, adOpenKeyset, adLockReadOnly
    If RS.RecordCount = 0 Then
        DBSYSTEM.Execute "INSERT INTO CFGASIENTOS VALUES(0,' ',0,' ',' ',' ',' ',' ',' ')"
    End If
End Sub
Public Function EXISTECONTA() As Boolean
On Error GoTo ERREXIS
    EXISTECONTA = False
    Dim RSAUX As New ADODB.Recordset
    Set RSAUX = New ADODB.Recordset
    RSAUX.Open "select * from sysdatabases where name ='BDWENCO'", CONECTARDBSQL("master"), adOpenKeyset, adLockReadOnly
    If RSAUX.RecordCount > 0 Then
        EXISTECONTA = True
    End If
    Exit Function
ERREXIS:
    EXISTECONTA = False
    Exit Function
End Function
Public Function CONECTARDBSQL(Optional BDNAME As String) As ADODB.Connection
On Error GoTo ERRNUM
    Set CONECTARDBSQL = New ADODB.Connection
    CONECTARDBSQL.CommandTimeout = 0
    CONECTARDBSQL.ConnectionTimeout = 0
    CONECTARDBSQL.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=SOPORTE;Password=SOPORTE;Initial Catalog=" & BDNAME & ";Data Source=" & VGL_SERVCONTA
    CONECTARDBSQL.Open
    Exit Function
ERRNUM:
    Set CONECTARDBSQL = Nothing
End Function
Public Sub SETCFGCONTA()
    Dim RSCFGCONTA As New ADODB.Recordset
    Set RSCFGCONTA = New ADODB.Recordset
    RSCFGCONTA.Open "CFGASIENTOS", DBSYSTEM, adOpenStatic, adLockReadOnly
    
    With REGSISTEMA
        .scTieneStConta = IIf(RSCFGCONTA("CHKCONTA"), True, False)
        .scRutaBDWenco = ""
        .scRutaEmpresaWenco = Trim(RSCFGCONTA("CODEMP"))
        .scTipoAnexo = Trim(RSCFGCONTA("TIPANEX"))
        .scSubdi = Trim(RSCFGCONTA("SUBDI"))
        .scCtaRedon = Trim(RSCFGCONTA("COMP"))
        .scCuenta = Trim(RSCFGCONTA("CUENTA"))
        .scCreaTrab = IIf(RSCFGCONTA("CHKCREATRAB"), True, False)
    End With
    If REGSISTEMA.scTieneStConta Then
        If EXISTECONTA Then
            REGSISTEMA.scNivelCta = DevuelveValor("SELECT EMP_NIVEL FROM EMPRESA WHERE EMP_CODIGO='" & RSCFGCONTA("CODEMP") & "'", CONECTARDBSQL("BDWENCO"))
        End If
      Else
        REGSISTEMA.scNivelCta = 0
    End If
End Sub
Public Function VERIFI_CONTA(TIPO As Integer, Optional AÑO As Long) As Boolean
    On Error GoTo ERRVER
    Dim RSAUX As New ADODB.Recordset
    Dim nomb As String
    Set RSAUX = Nothing
    VERIFI_CONTA = False
    Select Case TIPO
        Case 1
            nomb = REGSISTEMA.scRutaEmpresaWenco & "BDCONT" & Format(AÑO, "0000")
        Case 2
            nomb = REGSISTEMA.scRutaEmpresaWenco & "BDCONTABILIDAD"
    End Select
    RSAUX.Open "select * from sysdatabases where name='" & nomb & "'", CONECTARDBSQL("master"), adOpenKeyset, adLockReadOnly
    If RSAUX.RecordCount > 0 Then
        VERIFI_CONTA = True
      Else
       VERIFI_CONTA = False
    End If
    Exit Function
ERRVER:
    VERIFI_CONTA = False
End Function

Public Function Ultmes(FECHA As Date) As Integer
    Ultmes = Day(DateAdd("d", -1, DateAdd("m", 1, CDate("01/" & Month(FECHA) & "/" & Year(FECHA)))))
End Function
Sub ImpresionRptProc(cNombreReporte As String, PFormulas(), Param(), Optional ORDEN As String, Optional Titulo As String)
'Funcion creada por fernando cossio
Dim I As Integer
On Error GoTo X
    Screen.MousePointer = 11
    With MDIPrincipal.CryRptProc
        .Reset
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .WindowTitle = Titulo
        Call PropCrystal(MDIPrincipal.CryRptProc)
        .ReportFileName = REGSISTEMA.REPORTES & cNombreReporte
        .LogOnServer "pdssql.dll", VGL_SERVERREP, "" & VGL_BASE & "", "SOPORTE", "SOPORTE"
        .Connect = "DSN=" & VGL_SERVERREP & ";User=sa;PWD=;DSQ=" & VGL_BASE & ""
        .Formulas(0) = "@Emp='" & REGSISTEMA.EMPRESA & "'"
        .Formulas(1) = "@Ruc='" & REGSISTEMA.RUC & "'"
        .Formulas(2) = "@Dir='" & REGSISTEMA.DIRECCION & "'"
        If UBound(PFormulas) > 0 Then
            For I = 0 To UBound(PFormulas) - 1
                .Formulas(2 + I) = PFormulas(I)
            Next
        End If
        .DiscardSavedData = True
        If UBound(Param) > 0 Then
            For I = 0 To UBound(Param) - 1
                .StoredProcParam(I) = Param(I)
            Next
        End If
        If ORDEN <> "" Then Call CrystOrden(MDIPrincipal.CryRptProc, ORDEN)
        If .Status <> 2 Then .Action = 1
    End With
    Screen.MousePointer = 1
    Exit Sub
X:
  If ERR.Number = 9 Then Resume Next
  Screen.MousePointer = 1
  MsgBox "Error inesperado: " & ERR.Number & "  " & ERR.Description, vbExclamation
End Sub
Public Sub PropCrystal(ByRef CrystalRpt As CrystalReport)
    CrystalRpt.WindowShowCancelBtn = True
    CrystalRpt.WindowShowCloseBtn = True
    CrystalRpt.WindowShowExportBtn = True
    CrystalRpt.WindowShowGroupTree = True
    CrystalRpt.WindowShowNavigationCtls = True
    CrystalRpt.WindowShowPrintBtn = True
    CrystalRpt.WindowShowPrintSetupBtn = True
    CrystalRpt.WindowShowProgressCtls = True
    CrystalRpt.WindowShowSearchBtn = True
    CrystalRpt.WindowShowZoomCtl = True
End Sub
Private Sub CrystOrden(ByRef cry As CrystalReport, CAD As String)
Dim POS As Integer, cadaux As String, I As Integer
Dim VALOR As String
    Do While True
        POS = InStr(1, CAD, ",", vbTextCompare)
        I = 0
        If POS = 0 Then Exit Do
        VALOR = Left(CAD, POS - 1)
        cry.SortFields(I) = VALOR
        I = I + 1
        CAD = Right(CAD, (Len(CAD) - POS))
    Loop
End Sub
