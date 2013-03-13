VERSION 5.00
Object = "{C5FF36B7-A67E-11D3-9D9D-E6F193F7F854}#9.0#0"; "APLICTXT.OCX"
Begin VB.Form frAddEmp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nueva Empresa"
   ClientHeight    =   4770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5505
   Icon            =   "frAddEmp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4770
   ScaleWidth      =   5505
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmCancela 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   360
      Left            =   4005
      TabIndex        =   5
      Top             =   4395
      Width           =   1320
   End
   Begin VB.CommandButton cmAcepta 
      Caption         =   "&Aceptar"
      Height          =   360
      Left            =   4020
      TabIndex        =   4
      Top             =   4005
      Width           =   1320
   End
   Begin VB.Frame Frame1 
      Caption         =   "Información General"
      Height          =   1980
      Left            =   120
      TabIndex        =   0
      Top             =   105
      Width           =   5265
      Begin AplisetControlText.Aplitext xNombre 
         Height          =   315
         Left            =   1575
         TabIndex        =   2
         Top             =   795
         Width           =   3510
         _ExtentX        =   6191
         _ExtentY        =   556
         MaxLength       =   50
         Text            =   ""
      End
      Begin AplisetControlText.Aplitext xRuc 
         Height          =   300
         Left            =   1575
         TabIndex        =   1
         Top             =   405
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   529
         MaxLength       =   11
         Text            =   ""
         SinBlancos      =   -1  'True
         TipoDato        =   "N"
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   180
         Picture         =   "frAddEmp.frx":000C
         Top             =   1335
         Width           =   480
      End
      Begin VB.Label Label1 
         Caption         =   $"frAddEmp.frx":08D6
         Height          =   615
         Index           =   2
         Left            =   750
         TabIndex        =   8
         Top             =   1290
         Width           =   4350
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Razón Social"
         Height          =   195
         Index           =   1
         Left            =   150
         TabIndex        =   7
         Top             =   855
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Número de R.U.C."
         Height          =   195
         Index           =   0
         Left            =   165
         TabIndex        =   6
         Top             =   458
         Width           =   1305
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Directorio de Trabajo"
      Height          =   1695
      Left            =   120
      TabIndex        =   9
      Top             =   2160
      Width           =   5265
      Begin AplisetControlText.Aplitext NuevoDir 
         Height          =   285
         Left            =   3180
         TabIndex        =   3
         Top             =   765
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   503
         MaxLength       =   10
         Text            =   ""
         SinBlancos      =   -1  'True
         TipoCodigo      =   -1  'True
      End
      Begin VB.Image IMAG 
         Height          =   480
         Left            =   165
         Picture         =   "frAddEmp.frx":0972
         Top             =   1050
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nombre de Base de Datos SQL"
         Height          =   195
         Index           =   4
         Left            =   780
         TabIndex        =   13
         Top             =   825
         Width           =   2235
      End
      Begin VB.Label Label1 
         Caption         =   "Directorio donde se almacenan los datos del sistema de la empresa a crear. "
         Height          =   405
         Index           =   5
         Left            =   795
         TabIndex        =   12
         Top             =   1200
         Width           =   4245
      End
      Begin VB.Label xDir 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "No Seleccionada"
         Height          =   285
         Left            =   780
         TabIndex        =   11
         Top             =   435
         Width           =   4320
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ruta"
         Height          =   195
         Index           =   3
         Left            =   165
         TabIndex        =   10
         Top             =   465
         Width           =   345
      End
   End
   Begin VB.Label Label2 
      Height          =   270
      Left            =   225
      TabIndex        =   14
      Top             =   4380
      Width           =   3645
   End
End
Attribute VB_Name = "frAddEmp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmAcepta_Click()
                If Trim(xNombre.Text) = "" Then
                    MsgBox "El NOMBRE de la empresa no es valida", vbCritical
                    xNombre.SetFocus
                    Exit Sub
                End If
                If UCase(Left(xDir.Caption, 2)) = "A:" Or UCase(Left(xDir.Caption, 2)) = "B:" Then
                    MsgBox "No se permite grabar sobre una unidad de disco flexible", vbCritical
                    Exit Sub
                End If
Dim EX As Integer
    If VPTRASPRM = "NUEVA" Then
        EX = 0
        'VERIFICANDO SI YA EXISTE EL NRO DE RUC
        Dim RS As New ADODB.Recordset
        RS.Open "SELECT * FROM EMPRESAS WHERE RUC = '" & Trim(xRuc.Text) & "'", DBSTARPLAN, adOpenStatic, adLockOptimistic
        If RS.RecordCount Then
            MsgBox "El Numero de RUC ya existe en otra Empresa", vbInformation, "Información"
            xRuc.SetFocus
            Exit Sub
        End If
        'VERIFICANDO SI YA ESTA LA EMPRESA INGRESADA
        DBSTARPLAN.Execute "UPDATE EMPRESAS SET ACTIVO=ACTIVO WHERE RUC='" & xRuc.Text & "'", EX
                If EX > 0 Then
                    MsgBox "El número de RUC ya existe, deberá cambiarlo para poder continuar", vbCritical
                    xRuc.SetFocus
                    Exit Sub
                End If
                DBSTARPLAN.Execute "UPDATE EMPRESAS SET ACTIVO=ACTIVO WHERE NOMBRE='" & xNombre.Text & "'", EX
                If EX > 0 Then
                    MsgBox "El NOMBRE de la empresa ya existe, deberá cambiarlo para poder continuar", vbCritical
                    xNombre.SetFocus
                    Exit Sub
                End If
                If NuevoDir.Text = "" Then
                    MsgBox "El NOMBRE de la Base de Datos no puede ser una cadena vacia", vbCritical
                    NuevoDir.SetFocus
                    Exit Sub
                End If
                If UCase(Dir$(xDir.Caption, vbDirectory)) = "" Then
                    MsgBox "El directorio especificado en la configuración no existe", vbInformation
                    Exit Sub
                End If
                If UCase(Dir$(xDir.Caption & "\" & NuevoDir.Text & ".MDF", vbDirectory)) = UCase(NuevoDir.Text & ".DMF") Then
                    MsgBox "El NOMBRE de la Base de Datos ya existe", vbInformation
                    Exit Sub
                End If
                
                MkDir (xDir.Caption & "\" & NuevoDir.Text)
                MkDir (xDir.Caption & "\" & NuevoDir.Text & "\FOTOS")
                
                Dim CAD1 As String, CAD2 As String
                CAD1 = App.PATH & "\PLANILLA.SQL"
                
                If UCase(Dir(App.PATH & "\PLANILLA.SQL")) <> "PLANILLA.SQL" Then
                    MsgBox "Ud. no presenta derecho de creación de empresa", vbInformation
                Else
                    Label2.Caption = "Creando Base de Datos ..."
                    Label2.Refresh
                    Screen.MousePointer = vbHourglass
                        DBSTARPLAN.Execute "EXECUTE [" & VGL_SERVER & "].[STARPLAN].dbo.CREA_EMPRESA '" & UCase(Me.NuevoDir.Text) & "','" & UNIDADLOGICA & "\" & NuevoDir.Text & "\','" & CStr(Year(Date)) & "'"
                        DBSTARPLAN.Execute "INSERT INTO EMPRESAS (RUC, NOMBRE, DIRMASTER, DIRALMACEN, DIRREPORTS, ACTIVO, FECHACREACION) VALUES ('" & xRuc.Text & "','" & xNombre.Text & "','" & xDir.Caption & "\" & NuevoDir.Text & "','" & NuevoDir.Text & "','NONE', 0, " & DateSQL(Date) & ")"
                        Label2.Caption = "Actualizando información de la Base de Datos..."
                        Label2.Refresh
                        'CONEXION A LA NUEVA BASE PARA INSERTAR REGISTROS PREDETERMINADOS
                        Dim CNX_DATOS As New ADODB.Connection
                        'BASE DE DATOS NUEVA
                        With CNX_DATOS
                                .CursorLocation = adUseClient
                                .CommandTimeout = 0
                                If VGL_INTEGRWNT Then
                                    .ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=" & UCase(Me.NuevoDir.Text) & ""
                                 Else
                                    .ConnectionString = "PROVIDER=SQLOLEDB.1;PERSIST SECURITY INFO=FALSE;USER ID=SOPORTE;PASSWORD=SOPORTE;INITIAL CATALOG=" & UCase(Me.NuevoDir.Text) & ";DATA SOURCE=" & VGL_SERVER
                                End If
                                .Open
                        End With
                        CNX_DATOS.Execute "INSERT INTO EMPRESA (DIRMASTER, DIRALMACEN) VALUES('" & xDir.Caption & "\" & NuevoDir.Text & "', '" & xDir.Caption & "')"
                        'ACTUALIZACION DE VISTAS
                        '-----------------------
                        If IsEmpty(DevuelveValor("SELECT * FROM SYSOBJECTS WHERE NAME ='TRABBOLETEAR'", CNX_DATOS)) Then
                        CNX_DATOS.Execute "CREATE VIEW TRABBOLETEAR AS SELECT TRABAJADORES.CODTRAB, [TRABAJADORES].[APEPAT] + ' ' + LTrim([TRABAJADORES].[APEMAT]) + ' ' + LTrim([TRABAJADORES].[NOMBRE]) AS NOMBRES, TRABAJADORES.CCOSTO, TRABAJADORES.BASICO, AFPS.NOMBRE AS NOMAFP, AFPS.APOROBLI, AFPS.SEGURO, AFPS.TOPESEGURO, AFPS.COMISIONRA, TRABAJADORES.MESDEVENGUE, TRABAJADORES.ASIGFAM, CENTROSAR.NOMBRE AS NOMSCTR, CENTROSAR.TASA, TRABAJADORES.DEPARTAMENTO, TRABAJADORES.FONDOPENS, TRABAJADORES.AREA, TRABAJADORES.UBIGEO, TRABAJADORES.SEXO, TRABAJADORES.TIPOTRAB, TRABAJADORES.FECHAING, TRABAJADORES.SITUACIÓN, TRABAJADORES.CARGO, TRABAJADORES.BANCO, TRABAJADORES.ESSALUDVIDA, TRABAJADORES.RUCEPS, TRABAJADORES.NOPDT, TRABAJADORES.OPCION01, TRABAJADORES.OPCION02, TRABAJADORES.OPCIONA, TRABAJADORES.OPCIONB,TRABAJADORES.NOCALCULO " & _
                                    "FROM AFPS INNER JOIN (TRABAJADORES INNER JOIN CENTROSAR ON TRABAJADORES.CODSCTR = CENTROSAR.CODCAR) ON AFPS.CODAFP = TRABAJADORES.FONDOPENS "
                        End If
                        If IsEmpty(DevuelveValor("SELECT * FROM SYSOBJECTS WHERE NAME ='VWTRAB2'", CNX_DATOS)) Then
                            CNX_DATOS.Execute "CREATE VIEW VWTRAB2 AS SELECT TRABAJADORES.CODTRAB, LTRIM ([TRABAJADORES].[APEPAT])+ ' ' + LTRIM([TRABAJADORES].[APEMAT]) +' ' + LTRIM([TRABAJADORES].[NOMBRE]) AS NOMBRES,CCOSTOS.NOMBRE AS CENTRO, CCOSTOS.CODCCOSTO, TRABAJADORES.DOCIDEN,TRABAJADORES.FECHAING, TRABAJADORES.DEPARTAMENTO, TRABAJADORES.CARGO,TRABAJADORES.BASICO, TRABAJADORES.NUMFICHA, TRABAJADORES.CARNETSEG,TRABAJADORES.FONDOPENS, TRABAJADORES.ASIGFAM, TRABAJADORES.FECHACESE,TRABAJADORES.CODIGOALT, TRABAJADORES.TIPOTRAB, TRABAJADORES.SITUACIÓN,TRABAJADORES.CODSCTR, TRABAJADORES.RUCEPS FROM CCOSTOS INNER JOIN TRABAJADORES ON CCOSTOS.CODCCOSTO = TRABAJADORES.CCOSTO "
                        End If
                        If IsEmpty(DevuelveValor("SELECT * FROM SYSOBJECTS WHERE NAME ='VWTRABACTIVO'", CNX_DATOS)) Then
                            CNX_DATOS.Execute "CREATE VIEW VWTRABACTIVO AS SELECT TRABAJADORES.CODTRAB, LTRIM([TRABAJADORES].[APEPAT]) + ' ' + LTRIM([TRABAJADORES].[APEMAT]) + ' ' + LTRIM([TRABAJADORES].[NOMBRE]) AS NOMBRES, CCOSTOS.NOMBRE AS CCOSTO FROM CCOSTOS INNER JOIN TRABAJADORES ON CCOSTOS.CODCCOSTO = TRABAJADORES.CCOSTO WHERE (((TRABAJADORES.SITUACIÓN) <> '2'))"
                            End If
                        If IsEmpty(DevuelveValor("SELECT * FROM SYSOBJECTS WHERE NAME ='VWTRABAJ'", CNX_DATOS)) Then
                            CNX_DATOS.Execute "CREATE VIEW VWTRABAJ AS SELECT TRABAJADORES.CODTRAB, LTRIM ([TRABAJADORES].[APEPAT]) + ' ' + LTRIM([TRABAJADORES].[APEMAT]) + ' ' + LTRIM([TRABAJADORES].[NOMBRE]) AS NOMBRES, CCOSTOS.NOMBRE AS CENTRO, CCOSTOS.CODCCOSTO, AREASTRAB.NOMBRE AS NOMBREAREA, AREASTRAB.CODCCOSTO AS CODAREA, TRABAJADORES.DOCIDEN, TRABAJADORES.FECHAING, TRABAJADORES.DEPARTAMENTO, TRABAJADORES.CARGO, TRABAJADORES.BASICO, TRABAJADORES.NUMFICHA, TRABAJADORES.CARNETSEG, TRABAJADORES.FONDOPENS, TRABAJADORES.ASIGFAM, TRABAJADORES.FECHACESE, TRABAJADORES.CODIGOALT, TRABAJADORES.TIPOTRAB, TRABAJADORES.SITUACIÓN, TRABAJADORES.CODSCTR, TRABAJADORES.RUCEPS, AFPS.NOMBRE AS NOMBREAFP, TRABAJADORES.BANCO, TRABAJADORES.CTABANCO, TRABAJADORES.TIPDOC, TRABAJADORES.NOPDT AS NOPDT, TRABAJADORES.NOCALCULO FROM AREASTRAB INNER JOIN (CCOSTOS INNER JOIN (AFPS INNER JOIN TRABAJADORES ON AFPS.CODAFP = TRABAJADORES.FONDOPENS) ON CCOSTOS.CODCCOSTO = TRABAJADORES.CCOSTO) ON AREASTRAB.CODCCOSTO = TRABAJADORES.AREA "
                        End If
                        If IsEmpty(DevuelveValor("SELECT * FROM SYSOBJECTS WHERE NAME ='VWTRABAJGEN'", CNX_DATOS)) Then
                            CNX_DATOS.Execute "CREATE VIEW VWTRABAJGEN AS SELECT TRABAJADORES.CODTRAB, LTRIM ([TRABAJADORES].[APEPAT])+ ' ' + LTRIM([TRABAJADORES].[APEMAT])+ ' ' + LTRIM([TRABAJADORES].[NOMBRE]) AS NOMBRES,CCOSTOS.NOMBRE AS CENTRO, CCOSTOS.CODCCOSTO, AREASTRAB.NOMBRE AS NOMBREAREA, AREASTRAB.CODCCOSTO AS CODAREA,TRABAJADORES.DOCIDEN, TRABAJADORES.FECHAING,TRABAJADORES.DEPARTAMENTO, TRABAJADORES.CARGO,TRABAJADORES.BASICO, TRABAJADORES.NUMFICHA,TRABAJADORES.CARNETSEG, TRABAJADORES.FONDOPENS,TRABAJADORES.ASIGFAM, TRABAJADORES.FECHACESE,TRABAJADORES.CODIGOALT, TRABAJADORES.TIPOTRAB,TRABAJADORES.SITUACIÓN, TRABAJADORES.CODSCTR,TRABAJADORES.RUCEPS FROM AREASTRAB INNER JOIN (CCOSTOS INNER JOIN TRABAJADORES ON CCOSTOS.CODCCOSTO = TRABAJADORES.CCOSTO) ON AREASTRAB.CODCCOSTO = TRABAJADORES.AREA"
                        End If
                        If IsEmpty(DevuelveValor("SELECT * FROM SYSOBJECTS WHERE NAME ='VWUBIGEO'", CNX_DATOS)) Then
                            CNX_DATOS.Execute "CREATE VIEW VWUBIGEO AS SELECT UBIDIST.CODIGO, LTRIM([UBIDIST].[NOMBRE]) + ' - ' + LTRIM([UBIPROV].[NOMBRE]) + ' - ' + LTRIM([UBIDEP].[NOMBRE]) AS LUGAR FROM UBIDIST INNER JOIN (UBIPROV INNER JOIN UBIDEP ON UBIPROV.CODDEP = UBIDEP.CODIGO) ON UBIDIST.CODPROV = UBIPROV.CODIGO "
                        End If
                        If IsEmpty(DevuelveValor("SELECT * FROM SYSOBJECTS WHERE NAME ='ESTUDIOS_TRABAJADOR'", CNX_DATOS)) Then
                            CNX_DATOS.Execute "CREATE VIEW ESTUDIOS_TRABAJADOR AS SELECT TRABAJADORES.APEPAT, TRABAJADORES.APEMAT, TRABAJADORES.NOMBRE, ESTUDIOS.TIPOEST_DESCRIP, ESTUDIOS.EST_CARRERA, ESTUDIOS.DESCESTUDIOS, ESTUDIOS.TIPO_CESTUDIO, ESTUDIOS.EST_GRADO_OBT, ESTUDIOS.EST_FINI, ESTUDIOS.EST_FFIN, ESTUDIOS.EST_NIVEL, ESTUDIOS.EST_ADIC, ESTUDIOS.PAGEMP, ESTUDIOS.COD_TRAB, ESTUDIOS.CODIGO, CAST(TRABAJADORES.SITUACIÓN AS INT)AS SITUACIÓN FROM TRABAJADORES INNER JOIN ESTUDIOS ON TRABAJADORES.CODTRAB = ESTUDIOS.COD_TRAB"
                        End If
                        If IsEmpty(DevuelveValor("SELECT * FROM SYSOBJECTS WHERE NAME ='EVENTOS_TRABAJADOR'", CNX_DATOS)) Then
                            CNX_DATOS.Execute "CREATE VIEW EVENTOS_TRABAJADOR AS SELECT TRABAJADORES.CODTRAB, TRABAJADORES.CODIGOALT, TRABAJADORES.APEPAT, TRABAJADORES.APEMAT, TRABAJADORES.NOMBRE, EVENTOS.CATEGORIA, EVENTOS.SUBCATEGORIA, EVENTOS.ESTADO, EVENTOS.FEC_INI, EVENTOS.FEC_FIN, EVENTOS.HOR_INI, EVENTOS.HOR_FIN, EVENTOS.ASUNTO, CAST(TRABAJADORES.SITUACIÓN AS INT)AS SITUACIÓN FROM TRABAJADORES INNER JOIN EVENTOS ON TRABAJADORES.CODTRAB = EVENTOS.COD_TRABAJADOR"
                        End If
                        If IsEmpty(DevuelveValor("SELECT * FROM SYSOBJECTS WHERE NAME ='IDIOMAS_TRABAJADOR'", CNX_DATOS)) Then
                            CNX_DATOS.Execute "CREATE VIEW IDIOMAS_TRABAJADOR AS SELECT TRABAJADORES.APEPAT, TRABAJADORES.APEMAT, TRABAJADORES.NOMBRE, IDIOMAS.IDIO_DESCRIP, IDIOMAS.IDI_NIVEL, IDIOMAS.COD_TRAB, CAST(TRABAJADORES.SITUACIÓN AS INT)AS SITUACIÓN FROM TRABAJADORES INNER JOIN IDIOMAS ON TRABAJADORES.CODTRAB = IDIOMAS.COD_TRAB"
                        End If
                        If IsEmpty(DevuelveValor("SELECT * FROM SYSOBJECTS WHERE NAME ='LABORAL_TRABAJADOR'", CNX_DATOS)) Then
                            CNX_DATOS.Execute "CREATE VIEW LABORAL_TRABAJADOR AS SELECT TRABAJADORES.APEPAT, TRABAJADORES.APEMAT, TRABAJADORES.NOMBRE, LABORAL.LAB_CEN_LABORAL, LABORAL.CARGO, LABORAL.GIRO, LABORAL.LAB_FINI, LABORAL.LAB_FFIN, LABORAL.LAB_SUELDO_EST_ANUAL, LABORAL.FUNCION, LABORAL.CATEGORIA, LABORAL.LAB_CONDICION, LABORAL.LAB_MOT_SALIDA, LABORAL.LAB_COMENTARIO, LABORAL.COD_TRAB, CAST(TRABAJADORES.SITUACIÓN AS INT)AS SITUACIÓN FROM TRABAJADORES INNER JOIN LABORAL ON TRABAJADORES.CODTRAB = LABORAL.COD_TRAB"
                        End If
                        If IsEmpty(DevuelveValor("SELECT * FROM SYSOBJECTS WHERE NAME ='TMP_EVENTOS_CONS'", CNX_DATOS)) Then
                            CNX_DATOS.Execute "CREATE VIEW TMP_EVENTOS_CONS AS SELECT TRABAJADORES.CODTRAB, EVENTOS.CATEGORIA, EVENTOS.SUBCATEGORIA, EVENTOS.FEC_INI, EVENTOS.FEC_FIN, EVENTOS.HOR_INI, EVENTOS.HOR_FIN, EVENTOS.ASUNTO, EVENTOS.ESTADO FROM TRABAJADORES INNER JOIN EVENTOS ON TRABAJADORES.CODTRAB = EVENTOS.COD_TRABAJADOR"
                        End If
                        If IsEmpty(DevuelveValor("SELECT * FROM SYSOBJECTS WHERE NAME ='TRABAJADOR'", CNX_DATOS)) Then
                            CNX_DATOS.Execute "CREATE VIEW TRABAJADOR AS SELECT TRABAJADORES.CODTRAB, TRABAJADORES.APEPAT, TRABAJADORES.APEMAT, TRABAJADORES.NOMBRE, TRABAJADORES.SITUACIÓN From TRABAJADORES"
                        End If
                        
                            Dim SQLEXEC  As String
                            'GRABACION EN BLOQUES
                            SQLEXEC = "INSERT INTO BANCOS SELECT * FROM [" & VGL_SERVER & "].[STARPLAN].dbo.[BANCOS]"
                            CNX_DATOS.Execute SQLEXEC
                            SQLEXEC = "INSERT INTO DOCUMENTOS SELECT * FROM [" & VGL_SERVER & "].[STARPLAN].dbo.[DOCUMENTOS]"
                            CNX_DATOS.Execute SQLEXEC
                            SQLEXEC = "INSERT INTO UBIDEP SELECT * FROM [" & VGL_SERVER & "].[STARPLAN].dbo.[UBIDEP]"
                            CNX_DATOS.Execute SQLEXEC
                            SQLEXEC = "INSERT INTO UBIDIST SELECT * FROM [" & VGL_SERVER & "].[STARPLAN].dbo.[UBIDIST]"
                            CNX_DATOS.Execute SQLEXEC
                            SQLEXEC = "INSERT INTO UBIPROV SELECT * FROM [" & VGL_SERVER & "].[STARPLAN].dbo.[UBIPROV]"
                            CNX_DATOS.Execute SQLEXEC
                            SQLEXEC = "INSERT INTO COLUMPL SELECT * FROM [" & VGL_SERVER & "].[STARPLAN].dbo.[COLUMPL]"
                            CNX_DATOS.Execute SQLEXEC
                            SQLEXEC = "INSERT INTO CONCEPTOS SELECT * FROM [" & VGL_SERVER & "].[STARPLAN].dbo.[CONCEPTOS]"
                            CNX_DATOS.Execute SQLEXEC
                            SQLEXEC = "INSERT INTO BILLETES SELECT * FROM [" & VGL_SERVER & "].[STARPLAN].dbo.[BILLETES]"
                            CNX_DATOS.Execute SQLEXEC
                            SQLEXEC = "INSERT INTO TIPOSTRAB SELECT * FROM [" & VGL_SERVER & "].[STARPLAN].dbo.[TIPOSTRAB]"
                            CNX_DATOS.Execute SQLEXEC
                            SQLEXEC = "INSERT INTO CATEGORIA_EVENTOS SELECT * FROM [" & VGL_SERVER & "].[STARPLAN].dbo.[CATEGORIA_EVENTOS]"
                            CNX_DATOS.Execute SQLEXEC
                            SQLEXEC = "INSERT INTO DESCESTUDIOS SELECT * FROM [" & VGL_SERVER & "].[STARPLAN].dbo.[DESCESTUDIOS]"
                            CNX_DATOS.Execute SQLEXEC
                            SQLEXEC = "INSERT INTO ESTADO_EVENTO SELECT * FROM [" & VGL_SERVER & "].[STARPLAN].dbo.[ESTADO_EVENTO]"
                            CNX_DATOS.Execute SQLEXEC
                            SQLEXEC = "INSERT INTO GIROEMPRESA SELECT * FROM [" & VGL_SERVER & "].[STARPLAN].dbo.[GIROEMPRESA]"
                            CNX_DATOS.Execute SQLEXEC
                            SQLEXEC = "INSERT INTO SUBCATEGORIA SELECT * FROM [" & VGL_SERVER & "].[STARPLAN].dbo.[SUBCATEGORIA]"
                            CNX_DATOS.Execute SQLEXEC
                            SQLEXEC = "INSERT INTO TIPO_ESTUDIOS SELECT * FROM [" & VGL_SERVER & "].[STARPLAN].dbo.[TIPO_ESTUDIOS]"
                            CNX_DATOS.Execute SQLEXEC
                            SQLEXEC = "INSERT INTO TIPO_IDIOMAS SELECT * FROM [" & VGL_SERVER & "].[STARPLAN].dbo.[TIPO_IDIOMAS]"
                            CNX_DATOS.Execute SQLEXEC
                            SQLEXEC = "INSERT INTO AFPS SELECT * FROM [" & VGL_SERVER & "].[STARPLAN].dbo.[AFPS]"
                            CNX_DATOS.Execute SQLEXEC
                            SQLEXEC = "INSERT INTO CARGOS SELECT * FROM [" & VGL_SERVER & "].[STARPLAN].dbo.[CARGOS]"
                            CNX_DATOS.Execute SQLEXEC
                            SQLEXEC = "INSERT INTO TIPO_CENTROE SELECT * FROM [" & VGL_SERVER & "].[STARPLAN].dbo.[TIPO_CENTROE]"
                            CNX_DATOS.Execute SQLEXEC
                            'ActualizarSistema
                            Screen.MousePointer = vbNormal
                        Label2.Caption = ""
                        Label2.Refresh
                End If
    Else    'QUIERE DECIR QUE VPTRASPRM="EDITAR"
            DBSTARPLAN.Execute "UPDATE EMPRESAS SET RUC='" & xRuc.Text & "', NOMBRE='" & xNombre.Text & "', DIRMASTER='" & xDir.Caption & "' WHERE RUC='" & VPTRASPRM & "'"
    End If
    VPTAREA = "OK"
    Unload Me
End Sub

Private Sub CMCANCELA_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    xDir.Caption = REGSISTEMA.PATH
    Label2.Caption = ""
End Sub
Private Sub IMAG_Click()
    frSelDir.Show 1
    If VPTAREA <> "" Then xDir.Caption = VPTAREA
End Sub

Private Sub IMAGE2_Click()
    MsgBox "ENTERPRISE SOLUTIONS S.A."
End Sub


