VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C5FF36B7-A67E-11D3-9D9D-E6F193F7F854}#9.0#0"; "ApliCTxt.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frVerPlan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Visor de Planillas de Remuneraciones"
   ClientHeight    =   6090
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9300
   Icon            =   "frVerPlan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   9300
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "&Imprimir Formato"
      Height          =   390
      Left            =   7320
      TabIndex        =   14
      Top             =   4740
      Width           =   1875
   End
   Begin Crystal.CrystalReport Reporte 
      Left            =   345
      Top             =   2010
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cerrar"
      Height          =   390
      Left            =   7305
      TabIndex        =   5
      Top             =   5175
      Width           =   1875
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Imprimir"
      Height          =   390
      Left            =   7305
      TabIndex        =   4
      Top             =   4290
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.Frame Frame1 
      Caption         =   "Configuración de la Impresión"
      Height          =   2070
      Left            =   120
      TabIndex        =   3
      Top             =   3885
      Width           =   6990
      Begin AplisetControlText.Aplitext xTitulo 
         Height          =   300
         Index           =   0
         Left            =   1515
         TabIndex        =   7
         Top             =   450
         Width           =   5310
         _ExtentX        =   9366
         _ExtentY        =   529
         Text            =   ""
      End
      Begin AplisetControlText.Aplitext xTitulo 
         Height          =   300
         Index           =   1
         Left            =   1515
         TabIndex        =   11
         Top             =   820
         Width           =   5310
         _ExtentX        =   9366
         _ExtentY        =   529
         Text            =   ""
      End
      Begin AplisetControlText.Aplitext xTitulo 
         Height          =   300
         Index           =   2
         Left            =   1515
         TabIndex        =   12
         Top             =   1190
         Width           =   5310
         _ExtentX        =   9366
         _ExtentY        =   529
         Text            =   ""
      End
      Begin AplisetControlText.Aplitext xTitulo 
         Height          =   300
         Index           =   3
         Left            =   1515
         TabIndex        =   13
         Top             =   1560
         Width           =   5310
         _ExtentX        =   9366
         _ExtentY        =   529
         Text            =   ""
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Encabezado 03"
         Height          =   195
         Index           =   2
         Left            =   225
         TabIndex        =   10
         Top             =   1613
         Width           =   1125
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Encabezado 02"
         Height          =   195
         Index           =   1
         Left            =   210
         TabIndex        =   9
         Top             =   1244
         Width           =   1125
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Encabezado 01"
         Height          =   195
         Index           =   0
         Left            =   225
         TabIndex        =   8
         Top             =   877
         Width           =   1125
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Titulo del Informe"
         Height          =   195
         Left            =   195
         TabIndex        =   6
         Top             =   510
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmCfgVista 
      Caption         =   "&Configurar Vistas"
      Height          =   390
      Left            =   7305
      TabIndex        =   1
      ToolTipText     =   "Para agregar o quitar columnas de impresión"
      Top             =   3840
      Width           =   1875
   End
   Begin MSDataGridLib.DataGrid DataPlan 
      Height          =   3690
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   9150
      _ExtentX        =   16140
      _ExtentY        =   6509
      _Version        =   393216
      HeadLines       =   1
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
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label iSuma 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Ver Totales"
      Height          =   195
      Left            =   7770
      TabIndex        =   2
      Top             =   5685
      Width           =   1380
   End
   Begin VB.Image xSuma 
      Height          =   240
      Left            =   8025
      Picture         =   "frVerPlan.frx":08CA
      Top             =   5685
      Width           =   240
   End
End
Attribute VB_Name = "frVerPlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RSPLAN As ADODB.Recordset
Dim CadCnn As String
Dim CadenaAux As String

Private Sub cmCfgVista_Click()
    frCfgPl.Show 1
End Sub

Private Sub CmdImprimir_Click()
    Dim REG As Long
    Screen.MousePointer = 11
    Call CREARPLAN(REG)
    With Reporte
        .Reset
        '.LogOnServer "pdssql.dll", VGL_SERVERREP, "MARFICE_PP", "SOPORTE", "SOPORTE"
        .Connect = "DSN=" & VGL_SERVERREP & ";User=sa;PWD=;DSQ=planilla_pp"
        .ReportFileName = REGSISTEMA.REPORTES & "PLAN0066.RPT"
        .StoredProcParam(0) = " [##TMPCREPLAN" & VGL_COMPUTER & "] "
        .WindowTitle = "PLAN0066.RPT -" & DataPlan.Caption
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .WindowShowPrintBtn = True
        .WindowShowRefreshBtn = True
        .WindowShowSearchBtn = True
        .WindowShowPrintSetupBtn = True
        .Formulas(0) = "xCabeza='" & DataPlan.Caption & "'"
        .Formulas(1) = "xEmpresa='" & REGSISTEMA.EMPRESA & "'"
        .Formulas(2) = "xRuc='" & REGSISTEMA.RUC & "'"
        .Formulas(3) = "xReg='" & Str(REG) & "'"
        If .Status <> 2 Then .Action = 1
    End With
    Screen.MousePointer = 1
End Sub
Private Sub CREARPLAN(Optional ByRef REG As Long)
    Dim RSTRABPLAN As New ADODB.Recordset
    Dim RSAUX As New ADODB.Recordset
    Dim I As Integer, ORDEN As Integer
    Dim CONC As String
    
    If ExisteTablaAux(" [##TMPCREPLAN" & VGL_COMPUTER & "] ") Then DBSYSTEM.Execute "Drop Table  [##TMPCREPLAN" & VGL_COMPUTER & "] "
    DBSYSTEM.Execute "Create table  [##TMPCREPLAN" & VGL_COMPUTER & "]  (CODTRAB VARCHAR(8), NOMBRES VARCHAR(100), FECHAING DATETIME, CCOSTO VARCHAR(10),CARGO VARCHAR(30),BASICO  Numeric(20,2) , FONDOPENS VARCHAR(2), CODCONCEP VARCHAR(15),DESCONCEP VARCHAR(40), ORDEN INT, MONTO  Numeric(20,2) )"
    RSAUX.Open CadenaAux, DBSYSTEM, adOpenKeyset, adLockReadOnly
    REG = RSAUX.RecordCount
    RSTRABPLAN.Open " [##TMPCREPLAN" & VGL_COMPUTER & "] ", DBSYSTEM, adOpenKeyset, adLockOptimistic
    If RSAUX.RecordCount = 0 Then
        MsgBox "No existe ningún registro para imprimir la planilla"
        Exit Sub
    End If
    RSAUX.MoveFirst
    Do While Not RSAUX.EOF
        For I = 21 To RSAUX.Fields.Count - 1
            ORDEN = ORDEN + 1
            If DataPlan.Columns.Item(RSAUX.Fields(I).Name).Visible And _
            RSAUX.Fields(I).Value > 0 Then
                RSTRABPLAN.AddNew
                RSTRABPLAN!CODTRAB = RSAUX!CODTRAB
                RSTRABPLAN!NOMBRES = RSAUX!NOMBRES
                RSTRABPLAN!FECHAING = CDate(RSAUX!FECHAING)
                RSTRABPLAN!CCosto = RSAUX!CCosto
                RSTRABPLAN!CARGO = RSAUX!CARGO
                RSTRABPLAN!BASICO = RSAUX!BASICO
                RSTRABPLAN!FONDOPENS = RSAUX!FONDOPENS
                RSTRABPLAN!CODCONCEP = Trim(RSAUX.Fields(I).Name)
                CONC = DevuelveValor("Select NOMBRE From COLUMPL Where CODIGO='" & _
                                     Trim(RSAUX.Fields(I).Name) & "'", DBSYSTEM)
                If Trim(CONC) = "" Then
                    Select Case UCase(Trim(RSAUX.Fields(I).Name))
                        Case "OTROSINGR": CONC = "Otros INGRESOS"
                        Case "TOTING": CONC = "TOTAL INGRESOS"
                        Case "OTROSEGRE": CONC = "Otros Egresos"
                        Case "ADELANTO": CONC = "Adelanto"
                        Case "TOTEGR": CONC = "TOTAL Egresos"
                        Case "NETO", "NETOPAGO": CONC = "Neto"
                    End Select
                End If
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

Private Sub Command1_Click() 'Imprimir planilla
    If ExisteTablaAux(" [##PRTPLANILLA" & VGL_COMPUTER & "] ") Then DBSYSTEM.Execute "DROP TABLE  [##PRTPLANILLA" & VGL_COMPUTER & "] "
    DBSYSTEM.Execute "CREATE TABLE  [##PRTPLANILLA" & VGL_COMPUTER & "]  (FILA VARCHAR(5000))"
    RSPLAN.MoveFirst
    Dim X As Integer, CAD As String, XSTR As String, XNUMCHAR As Integer, xValor As Variant, XN As Byte, NCOUNT As Integer
    Dim RSAUX As ADODB.Recordset, CADTABLA As String
    Set RSAUX = New ADODB.Recordset
    Screen.MousePointer = 11
    
    'Proceso del encabezado de las planillas
    '---------------------------------------
    Set RSAUX = New ADODB.Recordset
    For X = 0 To DataPlan.Columns.Count - 1
        If DataPlan.Columns(X).Visible Then
            Select Case RSPLAN.Fields(DataPlan.Columns(X).Caption).Type
                    Case 135 'Si es por fecha
                        XNUMCHAR = 9
                    Case 6 'Si es numero simple
                        XNUMCHAR = 10
                    Case 200
                        XNUMCHAR = RSPLAN.Fields(DataPlan.Columns(X).Caption).DefinedSize + 1
                    Case 3
                        XNUMCHAR = 3
            End Select
            RSAUX.Open "SELECT NOMBRE FROM COLUMPL WHERE CODIGO='" & DataPlan.Columns(X).Caption & "'", DBSYSTEM, adOpenStatic
            If RSAUX.RecordCount = 0 Then
                CAD = DataPlan.Columns(X).Caption
            Else
                CAD = RSAUX!NOMBRE
            End If
            RSAUX.Close
            XSTR = XSTR & Left(CAD & String(XNUMCHAR, " "), XNUMCHAR)
        End If
    Next
    CAD = XSTR
    Set RSAUX = Nothing
    NCOUNT = 0
    
    'Generacion de la cadena de impresion
    '------------------------------------
    Do While Not RSPLAN.EOF
        XSTR = ""
        For X = 0 To DataPlan.Columns.Count - 1
            xValor = RSPLAN.Fields(DataPlan.Columns(X).Caption).Value
            If DataPlan.Columns(X).Visible Then
                Select Case RSPLAN.Fields(DataPlan.Columns(X).Caption).Type
                    Case 135 'Si es por fecha
                        If IsNull(xValor) Then xValor = "  /  /    "
                        XSTR = XSTR & " " & xValor
                    Case 6 'Si es numero simple
                        XSTR = XSTR & " " & Right("         " & Format$(xValor, "0.00"), 9)
                    Case 200
                        XN = RSPLAN.Fields(DataPlan.Columns(X).Caption).DefinedSize
                        XSTR = XSTR & " " & Left(xValor & String(XN, " "), XN)
                    Case 3
                        XSTR = XSTR & " " & Format(xValor, "00")
                    Case 11
                        
                    Case Else
                        MsgBox "Tipo no encontrado: " & RSPLAN.Fields(DataPlan.Columns(X).Caption).Type, vbCritical
                End Select
            End If
        Next
        NCOUNT = NCOUNT + 1
        If NCOUNT = 1 Then
            DBSYSTEM.Execute "INSERT INTO  [##PRTPLANILLA" & VGL_COMPUTER & "]  VALUES ('" & String(Len(XSTR), "_") & "')"
            'Incluir aqui la cabezar
            DBSYSTEM.Execute "INSERT INTO  [##PRTPLANILLA" & VGL_COMPUTER & "]  VALUES ('" & CAD & "')"
            DBSYSTEM.Execute "INSERT INTO  [##PRTPLANILLA" & VGL_COMPUTER & "]  VALUES ('" & String(Len(XSTR), "_") & "')"
        End If
        DBSYSTEM.Execute "INSERT INTO  [##PRTPLANILLA" & VGL_COMPUTER & "]  VALUES ('" & XSTR & "')"
        RSPLAN.MoveNext
    Loop
    CAD = String(Len(XSTR), "_")
    DBSYSTEM.Execute "INSERT INTO  [##PRTPLANILLA" & VGL_COMPUTER & "]  VALUES ('" & CAD & "')"
    
    'Proceso del Cálculo de TOTALes de la planilla
    '---------------------------------------------
    CADTABLA = REGSISTEMA.TABLAPLAN & " WHERE Mes=" & DateSQL(FechaMMAAAA(Right(frPlans.LPlans.SelectedItem.KEY, 6)))
    Set RSAUX = New ADODB.Recordset
    XSTR = ""
    For X = 0 To DataPlan.Columns.Count - 1
        If DataPlan.Columns(X).Visible Then
            Select Case RSPLAN.Fields(DataPlan.Columns(X).Caption).Type
                    Case 135 'Si es por fecha
                        XNUMCHAR = 9
                    Case 6 'Si es numero simple
                        XNUMCHAR = 10
                    Case 200
                        XNUMCHAR = RSPLAN.Fields(DataPlan.Columns(X).Caption).DefinedSize + 1
                    Case 3
                        XNUMCHAR = 3
            End Select
            If RSPLAN.Fields(DataPlan.Columns(X).Caption).Type = adSingle Then
                RSAUX.Open "SELECT SUM(" & DataPlan.Columns(X).Caption & ") as TOTAL FROM " & CADTABLA, DBSYSTEM, adOpenStatic
                If RSAUX.RecordCount = 0 Or IsNull(RSAUX!TOTAL) Then xValor = 0 Else xValor = RSAUX!TOTAL
                CAD = Right("         " & Format$(xValor, "0.00"), XNUMCHAR)
                RSAUX.Close
            Else
                CAD = String(XNUMCHAR, " ")
            End If
            XSTR = XSTR & CAD
        End If
    Next
    XSTR = "TOTAL General: " & Right(XSTR, Len(XSTR) - Len("TOTAL General: "))
    Set RSAUX = Nothing
    DBSYSTEM.Execute "INSERT INTO  [##PRTPLANILLA" & VGL_COMPUTER & "]  VALUES ('" & XSTR & "')"
    CAD = String(Len(XSTR), "_")
    DBSYSTEM.Execute "INSERT INTO  [##PRTPLANILLA" & VGL_COMPUTER & "]  VALUES ('" & CAD & "')"
    Screen.MousePointer = 1
    With Reporte
        .WindowTitle = "PLAN0029 " & DataPlan.Caption
        .WindowShowPrintBtn = True
        .WindowShowRefreshBtn = True
        .WindowShowSearchBtn = True
        .WindowShowPrintSetupBtn = True
        .ReportFileName = REGSISTEMA.REPORTES & "PLAN0029.RPT"
        .Connect = "DSN=" & VGL_SERVER & ";User=sa;PWD=;DSQ=" & REGSISTEMA.BASESQL & ""
        .StoredProcParam(0) = " [##PRTPLANILLA" & VGL_COMPUTER & "] "
        .Destination = crptToWindow
        .WindowState = crptMaximized
        If xTitulo(0).Text = "" Then
            .Formulas(0) = "xTitulo=' " & DataPlan.Caption & "'"
        Else
            .Formulas(0) = "xTitulo=' " & xTitulo(0).Text & "'"
        End If
        .Formulas(1) = "xCabeza1=' " & xTitulo(1).Text & "'"
        .Formulas(2) = "xCabeza2=' " & xTitulo(2).Text & "'"
        .Formulas(3) = "xCabeza3=' " & xTitulo(3).Text & "'"
        If .Status <> 2 Then .Action = 1
    End With
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub DataPlan_HeadClick(ByVal COLINDEX As Integer)
    CARGAPLAN CadCnn & " ORDER BY " & DataPlan.Columns(COLINDEX).Caption
End Sub

Public Sub CARGAPLAN(CadenaSQL As String)
    CadenaAux = CadenaSQL
    Dim X As Integer
    If CadCnn = "" Then CadCnn = CadenaSQL
    Set RSPLAN = New ADODB.Recordset
    RSPLAN.Open CadenaSQL, DBSYSTEM, adOpenStatic
    Set DataPlan.DataSource = RSPLAN
    For X = 0 To RSPLAN.Fields.Count - 1
        If RSPLAN.Fields(X).Type = adSingle Then
            DataPlan.Columns(RSPLAN.Fields(X).Name).NumberFormat = "##,##0.00 "
            DataPlan.Columns(RSPLAN.Fields(X).Name).Alignment = dbgRight
        End If
    Next
    If ExisteTabla("VERCOLPL") Then
        Dim RSAUX As New ADODB.Recordset
        RSAUX.Open "VERCOLPL", DBSYSTEM, adOpenStatic
        Do While Not RSAUX.EOF
            DataPlan.Columns(RSAUX!Codigo).Visible = IIf(RSAUX!VALOR = 1, True, False)
            RSAUX.MoveNext
        Loop
        Set RSAUX = Nothing
    End If
    DataPlan.Caption = "Planilla de Remuneraciones : " & VPTAREA
End Sub

Private Sub FORM_UNLOAD(CANCEL As Integer)
    Set RSPLAN = Nothing
    CadCnn = ""
End Sub

Private Sub iSuma_Click()
    xSuma_Click
End Sub

Private Sub xSuma_Click()
    VPTAREA = "Planillas"
    frSuma.Show 1
End Sub


