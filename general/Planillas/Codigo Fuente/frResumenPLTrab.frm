VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C5FF36B7-A67E-11D3-9D9D-E6F193F7F854}#9.0#0"; "ApliCTxt.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frResumenPLTrab 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Resumen de Planillas por Trabajador"
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9600
   Icon            =   "frResumenPLTrab.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   9600
   Begin VB.CommandButton cmQuitar 
      Caption         =   "&Quitar Columna"
      Height          =   315
      Left            =   6285
      TabIndex        =   21
      Top             =   6255
      Width           =   1440
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Cerrar"
      Height          =   315
      Left            =   7995
      TabIndex        =   20
      Top             =   6255
      Width           =   1440
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "&Imprimir Resumen"
      Height          =   315
      Left            =   4710
      TabIndex        =   19
      Top             =   6255
      Width           =   1440
   End
   Begin Crystal.CrystalReport rptBoletas 
      Left            =   4050
      Top             =   2940
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Boleta de Pago"
      Height          =   315
      Left            =   3135
      TabIndex        =   18
      Top             =   6255
      Width           =   1440
   End
   Begin AplisetControlText.Aplitext xCodFormula 
      Height          =   285
      Left            =   3735
      TabIndex        =   17
      Top             =   1470
      Visible         =   0   'False
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   503
      MaxLength       =   12
      Text            =   ""
      TipoCodigo      =   -1  'True
   End
   Begin AplisetControlText.Aplitext xConcepto 
      Height          =   285
      Left            =   5040
      TabIndex        =   16
      Top             =   1095
      Width           =   3915
      _ExtentX        =   6906
      _ExtentY        =   503
      Locked          =   -1  'True
      Text            =   ""
   End
   Begin VB.CommandButton cmFormula 
      Height          =   285
      Left            =   9045
      Picture         =   "frResumenPLTrab.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Ejecutar Fórmula"
      Top             =   1485
      Visible         =   0   'False
      Width           =   390
   End
   Begin AplisetControlText.Aplitext xFormula 
      Height          =   285
      Left            =   5055
      TabIndex        =   14
      Top             =   1470
      Visible         =   0   'False
      Width           =   3915
      _ExtentX        =   6906
      _ExtentY        =   503
      MaxLength       =   250
      Text            =   ""
   End
   Begin VB.CommandButton cmConceptos 
      Caption         =   "..."
      Height          =   285
      Left            =   9045
      TabIndex        =   12
      Top             =   1110
      Width           =   390
   End
   Begin MSDataGridLib.DataGrid xData 
      Height          =   4350
      Left            =   180
      TabIndex        =   10
      Top             =   1875
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   7673
      _Version        =   393216
      BackColor       =   15662071
      HeadLines       =   2
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
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
      Caption         =   "Resumen de Planillas"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   2
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.ComboBox xFiltro1 
      Height          =   315
      ItemData        =   "frResumenPLTrab.frx":0C0C
      Left            =   180
      List            =   "frResumenPLTrab.frx":0C19
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   6240
      Width           =   2865
   End
   Begin MSComCtl2.DTPicker xFechaIni 
      Height          =   300
      Left            =   765
      TabIndex        =   6
      Top             =   1065
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "MMM yyyy"
      Format          =   62062595
      CurrentDate     =   36875
   End
   Begin VB.Frame Frame1 
      Caption         =   "Información del Trabajdor"
      Height          =   810
      Left            =   105
      TabIndex        =   0
      Top             =   90
      Width           =   6525
      Begin AplisetControlText.Aplitext xTrab 
         Height          =   315
         Left            =   1080
         TabIndex        =   3
         Top             =   345
         Width           =   4830
         _ExtentX        =   8520
         _ExtentY        =   556
         Locked          =   -1  'True
         Text            =   ""
      End
      Begin VB.CommandButton cmMasInformacion 
         Caption         =   "+"
         Height          =   330
         Left            =   5970
         TabIndex        =   1
         ToolTipText     =   "Ver más datos del trabajador"
         Top             =   345
         Width           =   345
      End
      Begin VB.Label la 
         AutoSize        =   -1  'True
         Caption         =   "Trabajador"
         Height          =   195
         Left            =   135
         TabIndex        =   2
         Top             =   375
         Width           =   765
      End
   End
   Begin MSComCtl2.DTPicker xFechaFin 
      Height          =   300
      Left            =   765
      TabIndex        =   7
      Top             =   1470
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "MMM yyyy"
      Format          =   62062595
      CurrentDate     =   36875
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Calcular fórmula"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   2475
      TabIndex        =   13
      Top             =   1545
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Incluir concepto de remuneración"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   2475
      TabIndex        =   11
      Top             =   1140
      Width           =   2370
   End
   Begin VB.Image Image2 
      Height          =   240
      Left            =   9045
      Picture         =   "frResumenPLTrab.frx":0C73
      Top             =   285
      Width           =   240
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Resumen de Planillas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   7605
      TabIndex        =   9
      Top             =   300
      Width           =   1065
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   8805
      Picture         =   "frResumenPLTrab.frx":1AB5
      Top             =   225
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hasta"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   5
      Top             =   1515
      Width           =   420
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Desde"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   1110
      Width           =   465
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      Height          =   5640
      Left            =   105
      Top             =   975
      Width           =   9420
   End
End
Attribute VB_Name = "frResumenPLTrab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RSDATA As New ADODB.Recordset

Private Sub CMCONCEPTOS_Click()
    If RSDATA.RecordCount = 0 Then
        MsgBox "NO EXISTEN COINCIDENCIAS DEL TRABAJADOR EN LAS PLANILLAS ENTRE EL PERIODO SELECCIONADO", vbInformation
        Exit Sub
    End If
    Dim RSCONCEP As New ADODB.Recordset
    Dim xValor
    RSCONCEP.Open "SELECT CODIGO, NOMBRE,COMENTARIO FROM CONCEPTOS WHERE CODIGO<>'REDONDEO' ORDER BY NOMBRE", DBSYSTEM, adOpenKeyset, adLockOptimistic
    If RSCONCEP.RecordCount = 0 Then
        MsgBox "NO EXISTEN CONCEPTOS DE REMUNERACIONES EN LA BASE DE DATOS", vbInformation
        Set RSCONCEP = Nothing
        Exit Sub
    End If
    frmComun.CONECTAR RSCONCEP
    frmComun.Show 1
    If VGUTIL(1) <> "" Then
        If ExisteCampo(Trim(VGUTIL(1)), " [##TMPRESPLTRB" & VGL_COMPUTER & "] ", DBSYSTEM) Then
            MsgBox "ESTE CAMPO YA ESTA INCLUIDO EN LA CONSULTA", vbInformation
        Else
            xConcepto.Text = VGUTIL(2)
            xConcepto.Tag = VGUTIL(1)
            DBSYSTEM.Execute "ALTER TABLE  [##TMPRESPLTRB" & VGL_COMPUTER & "]  ADD " & VGUTIL(1) & "  Numeric(20,2)  " 'SINGLE"
            RSDATA.MoveFirst
            Do While Not RSDATA.EOF
                xValor = DevuelveValor("SELECT MONTO FROM MOV" & RSDATA!BOLETA & " WHERE INUMBOL=" & RSDATA!INUMBOL & " AND CONCEPTO='" & VGUTIL(1) & "'", DBSYSTEM)
                If IsNull(xValor) Or IsEmpty(xValor) Then xValor = 0
                DBSYSTEM.Execute "UPDATE  [##TMPRESPLTRB" & VGL_COMPUTER & "]  SET " & VGUTIL(1) & "=" & xValor & " WHERE CODIGO=" & RSDATA!Codigo & " AND INUMBOL=" & RSDATA!INUMBOL
                RSDATA.MoveNext
            Loop
            REFRESCARDATA
        End If
    End If
    Set RSCONCEP = Nothing
End Sub

Private Sub CMMASINFORMACION_Click()
    If xTrab.Tag = "" Then
        MsgBox "DEBE SELECCIONAR PRIMERO UN TRABAJADOR", vbInformation
        Exit Sub
    End If
    VPTAREA = xTrab.Tag
    CambiaPanelBD True
    Load frTrab
    CambiaPanelBD False
    frTrab.Show 1
End Sub

Private Sub CMQUITAR_CLICK()
    If xData.COL = -1 Then Exit Sub
    Select Case xData.Columns(xData.COL).DataField
        Case "CODIGO", "NOMBRE", "INUMBOL", "MES", "BOLETA"
            MsgBox "ESTA COLUMNA NO PUEDE SER ELIMINADA", vbInformation
        Case Else
            If MsgBox("SEGURO DE ELIMINAR LA COLUMNA " & xData.Columns(xData.COL).Caption, vbYesNo + vbQuestion) = vbNo Then Exit Sub
            DBSYSTEM.Execute "ALTER TABLE  [##TMPRESPLTRB" & VGL_COMPUTER & "]  DROP COLUMN " & xData.Columns(xData.COL).DataField
            REFRESCARDATA
    End Select
End Sub

Private Sub Command1_Click()
    If RSDATA.RecordCount = 0 Then
        MsgBox "NO SE HAN ENCONTRADO REGISTROS", vbInformation
        Exit Sub
    End If
    Dim XFILEBOL As String, xDir As String
    XFILEBOL = DevNomRep(Trim(REGSISTEMA.RUC), Trim(REGSISTEMA.USER), FILEBOLETA)
    If UCase(Dir$(REGSISTEMA.REPORTES & XFILEBOL)) <> UCase(XFILEBOL) Then
        MsgBox "NO SE HA ENCONTRADO EL REPORTE. ASIGNE CORRECTAMENTE EL NOMBRE DEL REPORTE DE BOLETAS DE REMUNERACIONES", vbInformation, "FALTA: " & XFILEBOL
        Exit Sub
    End If
    DBSTARPLAN.Execute "DELETE FROM RPTBOLETAS"
    Dim XBOOK As Variant
    If InStr(XFILEBOL, "X") > 0 Then
        On Error Resume Next
        If ExisteTablaAux(" [##RPTBOLCOLUMNAS" & VGL_COMPUTER & "] ") Then DBSYSTEM.Execute "DROP TABLE  [##RPTBOLCOLUMNAS" & VGL_COMPUTER & "] "
        DBSYSTEM.Execute "CREATE TABLE  [##RPTBOLCOLUMNAS" & VGL_COMPUTER & "]  (CODTRAB VARCHAR(8))"
        Dim TMPNOMBRE2 As String
        If Dir$(App.PATH & "\" & XFILEBOL) = XFILEBOL Then Kill App.PATH & "\" & XFILEBOL
        FileCopy REGSISTEMA.REPORTES & XFILEBOL, App.PATH & "\" & XFILEBOL
        TMPNOMBRE2 = Replace(XFILEBOL, "X", "0")
        If UCase(Dir$(REGSISTEMA.REPORTES & XFILEBOL)) <> UCase(XFILEBOL) Then
            MsgBox "NO EXISTE EL ARCHIVO AUXILIAR PARA LA IMPRESION DE ESTE REPORTE DE BOLETA DE REMUNERACIONES: FALTA " & TMPNOMBRE2, vbInformation
            Exit Sub
        End If
        If Dir$(App.PATH & "\" & TMPNOMBRE2) = TMPNOMBRE2 Then Kill App.PATH & "\" & TMPNOMBRE2
        FileCopy REGSISTEMA.REPORTES & TMPNOMBRE2, App.PATH & "\" & TMPNOMBRE2
    End If
    CambiaPanelBD True
    For Each XBOOK In xData.SelBookmarks
        xData.Bookmark = XBOOK
        CARGABOL
        If InStr(XFILEBOL, "X") > 0 Then DBSYSTEM.Execute "INSERT INTO  [##RPTBOLCOLUMNAS" & VGL_COMPUTER & "]  VALUES ('" & xTrab.Tag & "')"
    Next
    With rptBoletas
        'FRWAIT.SHOW 1
        .Reset
        xDir = DevuelveValor("SELECT DIRECCIÓN FROM EMPRESA", DBSYSTEM)
        .WindowTitle = "REPORTE DE BOLETAS DE REMUNERACIONES - RESUMEN DE PLANILLAS POR TRABAJADOR"
        If InStr(XFILEBOL, "X") > 0 Then
            .ReportFileName = App.PATH & "\" & XFILEBOL
        Else
            .ReportFileName = REGSISTEMA.REPORTES & XFILEBOL
            .Connect = "DSN=" & VGL_SERVER & ";User=sa;PWD=;DSQ=" & REGSISTEMA.BASESQL & ""
        End If
        .StoredProcParam(0) = "RPTBOLETAS"
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .Formulas(0) = "XEMPRESA='" & REGSISTEMA.EMPRESA & "'"
        .Formulas(1) = "XRUC='RUC N° " & REGSISTEMA.RUC & "'"
        .Formulas(2) = "XDIRECCION='" & xDir & "'"
        .WindowShowPrintBtn = True
        .WindowShowPrintSetupBtn = True
        '.WindowShowRefreshBtn = True
        .WindowShowSearchBtn = True
        CambiaPanelBD False
        If rptBoletas.Status <> 2 Then .Action = 1
    End With
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.TOP = 0
    Me.Left = 0
    xFechaFin.Value = Date
    xFechaIni.Value = Date
    xFechaIni.Day = 1
    xFechaIni.Value = DateAdd("M", -3, xFechaIni.Value)
    xFiltro1.ListIndex = 0
    If ExisteTablaAux(" [##TMPRESPLTRB" & VGL_COMPUTER & "] ") Then DBSYSTEM.Execute "DROP TABLE  [##TMPRESPLTRB" & VGL_COMPUTER & "] "
    DBSYSTEM.Execute "CREATE TABLE  [##TMPRESPLTRB" & VGL_COMPUTER & "]  (CODIGO INT, NOMBRE VARCHAR(50), INUMBOL INT, MES DATETIME, BOLETA VARCHAR(9), TOTING  Numeric(20,2) , TOTEGR  Numeric(20,2) , NETO  Numeric(20,2) )"
    REFRESCARDATA
End Sub

Private Sub XDATA_HEADCLICK(ByVal COLINDEX As Integer)
    RSDATA.Sort = xData.Columns(COLINDEX).DataField
End Sub

Private Sub XTRAB_DBLCLICK()
    Dim RSAUX As New ADODB.Recordset
    Select Case xFiltro1.ListIndex
        Case 0: RSAUX.Open "SELECT CODTRAB, NOMBRES FROM VWTRABAJ WHERE SITUACIÓN<'2'", DBSYSTEM, adOpenStatic, adLockReadOnly
        Case 1: RSAUX.Open "SELECT CODTRAB, NOMBRES FROM VWTRABAJ WHERE SITUACIÓN>='2'", DBSYSTEM, adOpenStatic, adLockReadOnly
        Case 2: RSAUX.Open "SELECT CODTRAB, NOMBRES FROM VWTRABAJ", DBSYSTEM, adOpenStatic, adLockReadOnly
    End Select
    If RSAUX.RecordCount = 0 Then
        MsgBox "NO SE HAN ENCONTRADO TRABAJADORES REGISTRADOS", vbInformation
        Set RSAUX = Nothing
        Exit Sub
    End If
    frmComun.CONECTAR RSAUX
    frmComun.Show 1
    If VGUTIL(1) <> "" Then
        xTrab.Text = RSAUX!NOMBRES
        xTrab.Tag = RSAUX!CODTRAB
        OBTENERVALORES
    End If
    Set RSAUX = Nothing
End Sub

Public Sub OBTENERVALORES()
    xFechaIni.Day = 1
    xFechaFin.Day = 1
    If xFechaIni.Value > xFechaFin.Value Then
        MsgBox "LA FECHA DE INICIO NO PUEDE SER MAYOR A LA FECHA FINAL", vbInformation
        Exit Sub
    End If
    Dim RSBOLS As New ADODB.Recordset
    Dim RSAUX As New ADODB.Recordset, XNUM As Integer
    If ExisteTablaAux(" [##TMPRESPLTRB" & VGL_COMPUTER & "] ") Then DBSYSTEM.Execute "DROP TABLE  [##TMPRESPLTRB" & VGL_COMPUTER & "] "
    DBSYSTEM.Execute "CREATE TABLE  [##TMPRESPLTRB" & VGL_COMPUTER & "]  (CODIGO INT, NOMBRE VARCHAR(50), INUMBOL INT, MES DATETIME, BOLETA VARCHAR(9), TOTING  Numeric(20,2) , TOTEGR  Numeric(20,2) , NETO  Numeric(20,2) )"
    RSAUX.Open "SELECT * FROM NOMBOL WHERE MES BETWEEN " & DateSQL(xFechaIni.Value) & " AND " & DateSQL(xFechaFin.Value) & " ORDER BY MES, FECHAINI", DBSYSTEM, adOpenStatic, adLockReadOnly
    Do While Not RSAUX.EOF
        If ExisteTabla("BOL" & Format(Month(RSAUX!MES), "00") & Year(RSAUX!MES)) Then
            RSBOLS.Open "SELECT * FROM BOL" & Format(Month(RSAUX!MES), "00") & Year(RSAUX!MES) & " WHERE CODNOMBOL=" & RSAUX!Codigo & " AND CODTRAB='" & xTrab.Tag & "'", DBSYSTEM, adOpenStatic, adLockReadOnly
            If RSBOLS.RecordCount <> 0 Then
                DBSYSTEM.Execute "INSERT INTO  [##TMPRESPLTRB" & VGL_COMPUTER & "]  VALUES (" & RSAUX!Codigo & ",'" & RSAUX!NOMBRE & "'," & RSBOLS!INUMBOL & "," & DateSQL(RSAUX!MES) & ",'" & Format(Month(RSAUX!MES), "00") & Year(RSAUX!MES) & "'," & RSBOLS!TOTING & "," & RSBOLS!TOTEGR & "," & (RSBOLS!TOTING - RSBOLS!TOTEGR) & ")"
            End If
            RSBOLS.Close
            End If
        RSAUX.MoveNext
    Loop
    Set RSBOLS = Nothing
    Set RSAUX = Nothing
    REFRESCARDATA
End Sub

Public Sub REFRESCARDATA()
    Set RSDATA = Nothing
    RSDATA.Open " [##TMPRESPLTRB" & VGL_COMPUTER & "] ", DBSYSTEM, adOpenStatic, adLockReadOnly
    Set xData.DataSource = RSDATA
    With xData
        .Columns("CODIGO").Visible = False
        .Columns("INUMBOL").Visible = False
        .Columns("MES").Visible = False
        .Columns("TOTING").Alignment = dbgRight
        .Columns("TOTEGR").Alignment = dbgRight
        .Columns("NETO").Alignment = dbgRight
        .Columns("TOTING").NumberFormat = "0.00 "
        .Columns("TOTEGR").NumberFormat = "0.00 "
        .Columns("NETO").NumberFormat = "0.00 "
        .Columns("NOMBRE").Width = 3450
        .Columns("TOTING").Width = 930
        .Columns("TOTEGR").Width = 930
        .Columns("NETO").Width = 930
        .Columns("BOLETA").Width = 930
        .Columns("TOTING").Caption = "TOTAL INGRESOS"
        .Columns("TOTEGR").Caption = "TOTAL EGRESOS"
        .Columns("NOMBRE").Caption = "PLANILLA DE REMUNERACIONES"
        If RSDATA.Fields.Count > 8 Then
            Dim XCOLS As Integer
            For XCOLS = 8 To RSDATA.Fields.Count - 1
                .Columns(Trim(RSDATA.Fields(XCOLS).Name)).Alignment = dbgRight
                .Columns(Trim(RSDATA.Fields(XCOLS).Name)).NumberFormat = "0.00 "
                .Columns(Trim(RSDATA.Fields(XCOLS).Name)).Width = 930
                .Columns(Trim(RSDATA.Fields(XCOLS).Name)).Caption = DevuelveValor("SELECT NOMBRE FROM CONCEPTOS WHERE CODIGO='" & Trim(RSDATA.Fields(XCOLS).Name) & "'", DBSYSTEM)
            Next
        End If
    End With
    xData.AllowUpdate = False
End Sub


Public Sub CARGABOL()
    Dim FMES As Date
    Dim ESVACACIONES As Boolean
    FMES = DevuelveValor("SELECT MES FROM NOMBOL WHERE CODIGO=" & RSDATA!Codigo, DBSYSTEM)
    ESVACACIONES = False
    If Not ExisteCampo("XREDONDEO", "RPTBOLETAS", DBSTARPLAN) Then
        DBSTARPLAN.Execute "ALTER TABLE RPTBOLETAS ADD XREDONDEO  Numeric(20,2) " 'SINGLE"
    End If
    If Not ExisteCampo("FIJO1", "RPTBOLETAS", DBSTARPLAN) Then
        DBSTARPLAN.Execute "ALTER TABLE RPTBOLETAS ADD FIJO1  Numeric(20,2) "
        DBSTARPLAN.Execute "ALTER TABLE RPTBOLETAS ADD FIJO2  Numeric(20,2) "
        DBSTARPLAN.Execute "ALTER TABLE RPTBOLETAS ADD FIJO3  Numeric(20,2) "
        DBSTARPLAN.Execute "ALTER TABLE RPTBOLETAS ADD FIJO4  Numeric(20,2) "
        DBSTARPLAN.Execute "ALTER TABLE RPTBOLETAS ADD FIJO5  Numeric(20,2) "
        DBSTARPLAN.Execute "ALTER TABLE RPTBOLETAS ADD FIJO7  Numeric(20,2) "
        DBSTARPLAN.Execute "ALTER TABLE RPTBOLETAS ADD FIJO8  Numeric(20,2) "
        DBSTARPLAN.Execute "ALTER TABLE RPTBOLETAS ADD FIJO9  Numeric(20,2) " 'SINGLE
    End If
    If Not ExisteCampo("CUSPP", "RPTBOLETAS", DBSTARPLAN) Then
        DBSTARPLAN.Execute "ALTER TABLE RPTBOLETAS ADD CUSPP VARCHAR(12)"  'TEXT
        DBSTARPLAN.Execute "ALTER TABLE RPTBOLETAS ADD AREA VARCHAR(25)"
    End If
    On Error GoTo ERRPRTBOL
    Dim RSAUX As New ADODB.Recordset
    Dim RSBOL As New ADODB.Recordset
    If ExisteTablaAux(" [##TMPTRANS" & VGL_COMPUTER & "] ") Then DBSYSTEM.Execute "DROP TABLE  [##TMPTRANS" & VGL_COMPUTER & "] "
    DBSYSTEM.Execute "CREATE TABLE  [##TMPTRANS" & VGL_COMPUTER & "]  (CODIGO VARCHAR(15), DESCRIPCION VARCHAR(30), VALOR  Numeric(20,2) , TIPO BIT, ENLACE VARCHAR(8), FILA INT, IMPRESIONFIJA BIT)"
                        'DBSYSTEM.EXECUTE "CREATE TABLE TMPTRANS (CODIGO TEXT(15), DESCRIPCION TEXT(30), VALOR SINGLE, TIPO BYTE, ENLACE TEXT(8), FILA SINGLE, IMPRESIONFIJA BYTE)"
    'JALAR LOS CONCEPTOS
    DBSYSTEM.Execute "INSERT INTO  [##TMPTRANS" & VGL_COMPUTER & "]  (CODIGO,DESCRIPCION,VALOR,TIPO,ENLACE,FILA, IMPRESIONFIJA) SELECT MOVICTA.CODMOV, DESCRIPCION, MONTO, 1 AS TIPO, ' ' AS ENLACE,11 AS FILA,0 AS IMPRESIONFIJA FROM MOVICTA, PAGOSCTA WHERE MOVICTA.CODMOV=PAGOSCTA.CODMOV AND MOVICTA.TIPOGRUPO=1 AND PAGOSCTA.CODTRAB='" & xTrab.Tag & "' AND CODNOMBOL=" & RSDATA!Codigo & " AND TIPOBOLETA='B'"
    'JALAR LOS OTROS INGRESOS (CUENTAS CORRIENTES)
    DBSYSTEM.Execute "INSERT INTO  [##TMPTRANS" & VGL_COMPUTER & "]  SELECT CONCEPTOS.CODIGO, CONCEPTOS.NOMBRE AS DESCRIPCION, MONTO AS VALOR, CONCEPTOS.TIPO, CONCEPTOS.ENLACE, FILA, IMPRESIONFIJA FROM MOV" & RSDATA!BOLETA & " MOV, CONCEPTOS WHERE MOV.CONCEPTO=CONCEPTOS.CODIGO AND INUMBOL=" & RSDATA!INUMBOL
    'JALAR LOS ADELANTOS DE PAGO
    DBSYSTEM.Execute "INSERT INTO  [##TMPTRANS" & VGL_COMPUTER & "]  SELECT CODIGO, '<ADELANTO DE PAGO>' AS DESCRIPCION,MONTO AS VALOR,2 AS TIPO,' ' AS ENLACE, 4 AS FILA,0 AS IMPRESIONFIJA FROM " & REGSISTEMA.TABLAADEL & " WHERE NOMBOL=" & RSDATA!Codigo & " AND CODTRAB='" & xTrab.Tag & "'"
    'JALAR LOS OTROS EGRESOS (CUENTAS CORRIENTES)
    DBSYSTEM.Execute "INSERT INTO  [##TMPTRANS" & VGL_COMPUTER & "]  (CODIGO,DESCRIPCION,VALOR,TIPO,ENLACE,FILA, IMPRESIONFIJA) SELECT MOVICTA.CODMOV, DESCRIPCION, MONTO, 2 AS TIPO, ' ' AS ENLACE,12 AS FILA,0 AS IMPRESIONFIJA FROM MOVICTA, PAGOSCTA WHERE MOVICTA.CODMOV=PAGOSCTA.CODMOV AND MOVICTA.TIPOGRUPO=2 AND PAGOSCTA.CODTRAB='" & xTrab.Tag & "' AND CODNOMBOL=" & RSDATA!Codigo & " AND TIPOBOLETA='B'"
    'COLOCANDO LA FECHA DE LA VACACIONES AL SISTEMA
    
    RSAUX.Open "SELECT * FROM VWTRABAJ WHERE CODTRAB='" & xTrab.Tag & "'", DBSYSTEM, adOpenStatic
    RSBOL.Open "RPTBOLETAS", DBSTARPLAN, adOpenKeyset, adLockOptimistic
    'ADICION DE LA BOLETA
    RSBOL.AddNew
    RSBOL!CODTRAB = xTrab.Tag
    RSBOL!NOMBRES = xTrab.Text
    RSBOL!CENTROCOSTO = RSAUX!CENTRO
    RSBOL!FECHAING = RSAUX!FECHAING
    RSBOL!PERIODO = RSDATA!NOMBRE
    RSBOL!TOTING = RSDATA!TOTING
    RSBOL!TOTEGR = RSDATA!TOTEGR
    RSBOL!XREDONDEO = DevuelveValor("SELECT XREDONDEO FROM BOL" & RSDATA!BOLETA & " WHERE INUMBOL=" & RSDATA!INUMBOL, DBSYSTEM)
    RSBOL!BASICO = DevuelveValor("SELECT BASICO FROM BOL" & RSDATA!BOLETA & " WHERE INUMBOL=" & RSDATA!INUMBOL, DBSYSTEM)
    RSBOL!AFP = "" & RSAUX!NOMBREAFP
    RSBOL!CARGO = RSAUX!CARGO
    RSBOL!FECHAING = RSAUX!FECHAING
    RSBOL!DOCUMENTO = "" & RSAUX!DOCIDEN
    RSBOL!CARNETSEG = RSAUX!CARNETSEG
    RSBOL!CUENTABANCO = "" & RSAUX!CTABANCO
    RSBOL!CUSPP = "" & DevuelveValor("SELECT CUSPP FROM TRABAJADORES WHERE CODTRAB='" & xTrab.Tag & "'", DBSYSTEM)
    RSBOL!AREA = "" & RSAUX!NOMBREAREA
    Dim IND1 As Byte, IND2 As Byte, IND3 As Byte, IND4 As Byte
    Dim CLASEPRT As Byte
    CLASEPRT = DevuelveValor("SELECT CLASEBOLETA FROM EMPRESA", DBSYSTEM)
    DevNomRep Trim(REGSISTEMA.RUC), Trim(REGSISTEMA.USER), FILEBOLETA, CLASEPRT
    Set RSAUX = Nothing
    
    Dim PERIODOVAC As String
    RSAUX.Open "SELECT * FROM HISTOVAC WHERE NOMBOL=" & RSDATA!Codigo & " AND CODTRAB='" & xTrab.Tag & "'", DBSYSTEM, adOpenStatic, adLockReadOnly
    If Not RSAUX.EOF Or RSAUX.RecordCount <> 0 Then
        RSBOL!FECHAVAC1 = RSAUX!FECHAINI
        RSBOL!FECHAVAC2 = RSAUX!FECHAFIN
        PERIODOVAC = RSAUX!PERIODO
        ESVACACIONES = True
    End If
    Set RSAUX = Nothing
    RSAUX.Open "SELECT * FROM  [##TMPTRANS" & VGL_COMPUTER & "]  ORDER BY TIPO, FILA", DBSYSTEM, adOpenStatic
    IND1 = 0
    IND2 = 0
    IND3 = 0
    IND4 = 0
    Dim XCT As String, XCN As String, XCONTEO, XTIPCNPT As Byte, XTOTALAPORT As Single
    XCONTEO = 0
    XCT = "C"
    XCN = "I"
    XTIPCNPT = 0
    XTOTALAPORT = 0
    Do While Not RSAUX.EOF
        If RSAUX!IMPRESIONFIJA = 1 Then
            RSBOL.Fields("FIJO" & RSAUX!FILA).Value = RSAUX!VALOR
        Else
            Select Case CLASEPRT
                Case 0
                    Select Case RSAUX!TIPO
                        Case 0: IND1 = IND1 + 1
                                RSBOL.Fields("C" & IND1).Value = RSAUX!DESCRIPCION
                                RSBOL.Fields("INF" & IND1).Value = RSAUX!VALOR
                        Case 1: IND1 = IND1 + 1
                                RSBOL.Fields("C" & IND1).Value = RSAUX!DESCRIPCION
                                If UCase(RSAUX!Codigo) = "REMUVAC" Then RSBOL.Fields("C" & IND1).Value = "REMU. VAC. " & PERIODOVAC
                                RSBOL.Fields("I" & IND1).Value = RSAUX!VALOR
                        Case 2: IND3 = IND3 + 1
                                RSBOL.Fields("R" & IND3).Value = RSAUX!DESCRIPCION
                                RSBOL.Fields("E" & IND3).Value = RSAUX!VALOR
                        Case 3: IND4 = IND4 + 1
                                RSBOL.Fields("G" & IND4).Value = RSAUX!DESCRIPCION
                                RSBOL.Fields("A" & IND4).Value = RSAUX!VALOR
                    End Select
                Case 1
                    Select Case RSAUX!TIPO
                        Case 0: IND1 = IND1 + 1
                                RSBOL.Fields("C" & IND1).Value = RSAUX!DESCRIPCION & "          (" & RSAUX!VALOR & ")"
                                RSBOL.Fields("INF" & IND1).Value = RSAUX!VALOR
                        Case 1: IND1 = IND1 + 1
                                RSBOL.Fields("C" & IND1).Value = RSAUX!DESCRIPCION
                                RSBOL.Fields("I" & IND1).Value = RSAUX!VALOR
                        Case 2: IND1 = IND1 + 1
                                RSBOL.Fields("C" & IND1).Value = RSAUX!DESCRIPCION
                                RSBOL.Fields("E" & IND1).Value = RSAUX!VALOR
                        Case 3: IND4 = IND4 + 1
                                RSBOL.Fields("G" & IND4).Value = RSAUX!DESCRIPCION
                                RSBOL.Fields("A" & IND4).Value = RSAUX!VALOR
                    End Select
                Case 2
                    If XCONTEO >= 21 Then
                        XCT = "R"
                        XCN = "E"
                        XCONTEO = XCONTEO - 20
                    End If
                    XCONTEO = XCONTEO + 1
                    If XTIPCNPT <> RSAUX!TIPO Then
                        XTIPCNPT = RSAUX!TIPO
                        Select Case XTIPCNPT
                            Case 2
                                If XCONTEO >= 21 Then
                                    XCT = "R"
                                    XCN = "E"
                                    XCONTEO = XCONTEO - 20
                                End If
                                RSBOL.Fields(XCT & XCONTEO).Value = "TOTAL INGRESOS"
                                RSBOL.Fields(XCN & XCONTEO).Value = RSDATA!TOTING
                                XCONTEO = XCONTEO + 2
                                If XCONTEO >= 21 Then
                                    XCT = "R"
                                    XCN = "E"
                                    XCONTEO = XCONTEO - 20
                                End If
                                RSBOL.Fields(XCT & XCONTEO).Value = "RETENCIONES Y DESCUENTOS"
                                RSBOL.Fields(XCN & XCONTEO).Value = 0
                                XCONTEO = XCONTEO + 1
                            Case 3
                                If XCONTEO >= 21 Then
                                    XCT = "R"
                                    XCN = "E"
                                    XCONTEO = XCONTEO - 20
                                End If
                                RSBOL.Fields(XCT & XCONTEO).Value = "TOTAL EGRESOS"
                                RSBOL.Fields(XCN & XCONTEO).Value = RSDATA!TOTEGR
                                XCONTEO = XCONTEO + 2
                                If XCONTEO >= 21 Then
                                    XCT = "R"
                                    XCN = "E"
                                    XCONTEO = XCONTEO - 20
                                End If
                                RSBOL.Fields(XCT & XCONTEO).Value = "NETO A PAGAR"
                                RSBOL.Fields(XCN & XCONTEO).Value = (RSDATA!TOTING - RSDATA!TOTEGR)
                                XCONTEO = XCONTEO + 2
                                If XCONTEO >= 21 Then
                                    XCT = "R"
                                    XCN = "E"
                                    XCONTEO = XCONTEO - 20
                                End If
                                If XCONTEO = 0 Then XCONTEO = XCONTEO + 1
                                RSBOL.Fields(XCT & XCONTEO).Value = "APORTACIONES DEL EMPLEADOR"
                                RSBOL.Fields(XCN & XCONTEO).Value = 0
                                XCONTEO = XCONTEO + 1
                        End Select
                    End If
                    If XCONTEO >= 21 Then
                        XCT = "R"
                        XCN = "E"
                        XCONTEO = XCONTEO - 20
                    End If
                    If RSAUX!TIPO = 3 Then XTOTALAPORT = XTOTALAPORT + RSAUX!VALOR
                    RSBOL.Fields(XCT & XCONTEO).Value = RSAUX!DESCRIPCION
                    RSBOL.Fields(XCN & XCONTEO).Value = RSAUX!VALOR
            End Select
        End If
        RSAUX.MoveNext
    Loop
    If CLASEPRT = 2 Then
        XCONTEO = XCONTEO + 1
        If XCONTEO >= 21 Then
            XCT = "R"
            XCN = "E"
            XCONTEO = XCONTEO - 20
        End If
        RSBOL.Fields(XCT & XCONTEO).Value = "TOTAL APORTACIONES"
        RSBOL.Fields(XCN & XCONTEO).Value = XTOTALAPORT
    End If
    RSBOL.Update
    If ESVACACIONES Then
        IND1 = IND1 + 1
        'FALTA CODIGO
    End If
    Set RSAUX = Nothing
    Set RSBOL = Nothing
    Exit Sub
ERRPRTBOL:
    Resume Next
End Sub

Private Sub CmdImprimir_Click()
    If RSDATA.RecordCount = 0 Then
        MsgBox "NO EXISTE INFORMACIÓN POR IMPRIMIR", vbInformation
        Exit Sub
    End If
    Dim REG As Long
    Screen.MousePointer = 11
    Call CREARPLAN(REG)
    With rptBoletas
        .Reset
        .WindowTitle = "PLAN0079.RPT -" & xData.Caption
        .ReportFileName = REGSISTEMA.REPORTES & "PLAN0079.RPT"
        .Connect = "DSN=" & VGL_SERVERREP & ";User=sa;PWD=;DSQ=planilla_pp"
        .StoredProcParam(0) = " [##TMPCREPLAN" & VGL_COMPUTER & "] "
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .WindowShowPrintBtn = True
        '.WindowShowRefreshBtn = True
        .WindowShowSearchBtn = True
        .WindowShowPrintSetupBtn = True
        .Formulas(0) = "XEMPRESA='" & REGSISTEMA.EMPRESA & "'"
        .Formulas(1) = "XRUC='" & REGSISTEMA.RUC & "'"
        .Formulas(2) = "XREG='" & Str(REG) & "'"
        .Formulas(3) = "XTRABAJADOR='" & xTrab.Text & "'"
        If .Status <> 2 Then .Action = 1
    End With
    Screen.MousePointer = 1
End Sub

Private Sub CREARPLAN(Optional ByRef REG As Long)
    Dim RSTRABPLAN As New ADODB.Recordset
    Dim RSAUX As New ADODB.Recordset
    Dim I As Integer, ORDEN As Integer
    Dim CONC As String
    Set RSAUX = New ADODB.Recordset
    Set RSTRABPLAN = New ADODB.Recordset
    If ExisteTablaAux("[##TMPCREPLAN" & VGL_COMPUTER & "]") Then DBSYSTEM.Execute "DROP TABLE  [##TMPCREPLAN" & VGL_COMPUTER & "] "
    DBSYSTEM.Execute "CREATE TABLE  [##TMPCREPLAN" & VGL_COMPUTER & "]  (MES DATETIME, CODTRAB VARCHAR(8),NOMBRES VARCHAR(50) ,CODCONCEP VARCHAR(15),DESCONCEP VARCHAR(40),ORDEN INT,MONTO  Numeric(20,2) )"
    RSAUX.Open " [##TMPRESPLTRB" & VGL_COMPUTER & "] ", DBSYSTEM, adOpenKeyset, adLockReadOnly
    REG = RSAUX.RecordCount
    RSTRABPLAN.Open " [##TMPCREPLAN" & VGL_COMPUTER & "] ", DBSYSTEM, adOpenKeyset, adLockOptimistic
    If RSAUX.RecordCount = 0 Then
        MsgBox "NO EXISTE NINGÚN REGISTRO PARA IMPRIMIR LA PLANILLA"
        Exit Sub
    End If
    RSAUX.MoveFirst
    Do While Not RSAUX.EOF
        For I = 5 To RSAUX.Fields.Count - 1
            ORDEN = ORDEN + 1
            If RSAUX.Fields(I).Value > 0 Then
                RSTRABPLAN.AddNew
                RSTRABPLAN!MES = RSAUX!MES
                RSTRABPLAN!CODTRAB = "PL" & RSAUX!Codigo
                RSTRABPLAN!NOMBRES = RSAUX!NOMBRE
                RSTRABPLAN!CODCONCEP = Trim(RSAUX.Fields(I).Name)
                Select Case RSAUX.Fields(I).Name
                    Case "TOTING"
                        CONC = "TOTAL INGRESOS"
                    Case "TOTEGR"
                        CONC = "TOTAL EGRESOS"
                    Case "NETO"
                        CONC = "NETO DEL PAGO"
                    Case Else
                        CONC = DevuelveValor("SELECT NOMBRE FROM CONCEPTOS WHERE CODIGO='" & Trim(RSAUX.Fields(I).Name) & "'", DBSYSTEM)
                End Select
                RSTRABPLAN!DESCONCEP = CONC
                RSTRABPLAN!ORDEN = ORDEN
                RSTRABPLAN!MONTO = RSAUX.Fields(I).Value
                RSTRABPLAN.Update
            End If
        Next
        ORDEN = 0
        RSAUX.MoveNext
    Loop
End Sub


