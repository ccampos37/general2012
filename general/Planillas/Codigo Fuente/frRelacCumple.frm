VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frRelacCumple 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Relación de Cumpleaños"
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6060
   Icon            =   "frRelacCumple.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   6060
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton Option4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Solo de la Empresa Activa"
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   240
      TabIndex        =   14
      Top             =   240
      Width           =   2175
   End
   Begin VB.OptionButton Option3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "De todas las Empresas "
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   3240
      TabIndex        =   13
      Top             =   240
      Width           =   1935
   End
   Begin Crystal.CrystalReport Reporte 
      Left            =   240
      Top             =   960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmCerrar 
      Caption         =   "&Cerrar"
      Height          =   390
      Left            =   4605
      TabIndex        =   11
      Top             =   3090
      Width           =   1335
   End
   Begin VB.CommandButton cmImprimir 
      Caption         =   "&Imprimir"
      Height          =   390
      Left            =   4605
      TabIndex        =   10
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CommandButton cmSelecTrab 
      Caption         =   "Seleccion (F5)"
      Height          =   990
      Left            =   135
      Picture         =   "frRelacCumple.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Seleccione aquí los trabajadores quienes particparan de la citación"
      Top             =   2520
      Width           =   945
   End
   Begin VB.Frame Frame1 
      Caption         =   "Configuración del Reporte"
      Height          =   1740
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   5820
      Begin MSComCtl2.DTPicker xAnno 
         Height          =   390
         Left            =   3390
         TabIndex        =   7
         Top             =   780
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   688
         _Version        =   393216
         CustomFormat    =   "yyyy"
         Format          =   61997059
         UpDown          =   -1  'True
         CurrentDate     =   36870
      End
      Begin MSComCtl2.DTPicker xMes 
         Height          =   390
         Left            =   3390
         TabIndex        =   5
         Top             =   1275
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   688
         _Version        =   393216
         CustomFormat    =   "MMMM"
         Format          =   61997059
         UpDown          =   -1  'True
         CurrentDate     =   36870
      End
      Begin VB.CheckBox xTodos 
         Caption         =   "Todos los meses"
         Height          =   240
         Left            =   705
         TabIndex        =   3
         Top             =   825
         Width           =   1575
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Reporte Artístico"
         Height          =   255
         Left            =   2805
         TabIndex        =   2
         Top             =   315
         Value           =   -1  'True
         Width           =   1635
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Reporte normal"
         Height          =   195
         Left            =   675
         TabIndex        =   1
         Top             =   345
         Width           =   1485
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "< Todos >"
         Height          =   255
         Left            =   3405
         TabIndex        =   12
         Top             =   1335
         Width           =   1455
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Año"
         Height          =   195
         Left            =   2625
         TabIndex        =   6
         Top             =   840
         Width           =   285
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Mes"
         Height          =   195
         Left            =   2625
         TabIndex        =   4
         Top             =   1305
         Width           =   300
      End
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "<< Trabajadores a Listar (Seleccione)"
      Height          =   195
      Left            =   1200
      TabIndex        =   8
      Top             =   3240
      Width           =   2640
   End
End
Attribute VB_Name = "frRelacCumple"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RS_GENERAL As ADODB.Recordset
Dim X As Integer
Private Sub cmCerrar_Click()
    Unload Me
End Sub
Private Sub CMIMPRIMIR_CLICK()
Dim RS_INSERT As ADODB.Recordset
Dim SQLINSERT As String, SQLDELETE As String
If Me.Option4.Value = True Then
        If ExisteTablaSQL(" [##_TMPCUMPLE" & VGL_COMPUTER & "] ", DBAUXCOM) Then DBSTARPLAN.Execute "DROP TABLE  [##_TMPCUMPLE" & VGL_COMPUTER & "] "
            DBSTARPLAN.Execute "CREATE TABLE  [##_TMPCUMPLE" & VGL_COMPUTER & "]  (CODTRAB VARCHAR(8), NOMBRES VARCHAR(20), APEPAT VARCHAR(20), APEMAT VARCHAR(20), FECHANAC DATETIME)"
        If xTodos.Value = 0 Then
            DBSTARPLAN.Execute "INSERT INTO  [##_TMPCUMPLE" & VGL_COMPUTER & "]  SELECT CODTRAB , NOMBRE, APEPAT, APEMAT, FECHANAC FROM " & REGSISTEMA.BASESQL & ".dbo.TRABAJADORES WHERE MONTH(FECHANAC)=" & xMes.Month & ""
        Else
            DBSTARPLAN.Execute "INSERT INTO  [##_TMPCUMPLE" & VGL_COMPUTER & "]  SELECT CODTRAB , NOMBRE, APEPAT, APEMAT, FECHANAC FROM " & REGSISTEMA.BASESQL & ".dbo.TRABAJADORES "
        End If
    With Reporte
        .Reset
        If Option1.Value Then
            .ReportFileName = REGSISTEMA.REPORTES & "PLNAS003.RPT"
            .WindowTitle = "PLNAS003- CUMPLEAÑOS"
        Else
            .ReportFileName = REGSISTEMA.REPORTES & "PLNAS006.RPT"
            .WindowTitle = "PLNAS006- CUMPLEAÑOS"
        End If
        .Connect = "DSN=" & VGL_SERVERREP & ";User=sa;PWD=;DSQ=planilla_pp"
        .Formulas(0) = "XEMPRESA='" & REGSISTEMA.EMPRESA & "'"
        .Formulas(1) = "XTITULO='" & IIf(xTodos.Value = 1, "Año" & Year(xAnno.Value), " Mes de " & AMESES(Month(xMes.Value)) & " de " & Year(xAnno.Value)) & "'"
        .StoredProcParam(0) = VGL_COMPUTER
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .WindowShowPrintBtn = True
        .WindowShowRefreshBtn = True
        .WindowShowSearchBtn = True
        .WindowShowPrintSetupBtn = True
        If .Status <> 2 Then .Action = 1
    End With
Else
    On Error Resume Next
    If ExisteTablaSQL(" [##_TMPCUMPLE" & VGL_COMPUTER & "] ", DBSYSTEM) Then DBSTARPLAN.Execute "DROP TABLE  [##_TMPCUMPLE" & VGL_COMPUTER & "] "
    DBSTARPLAN.Execute "CREATE TABLE  [##_TMPCUMPLE" & VGL_COMPUTER & "]  (CODTRAB VARCHAR(8), NOMBRES VARCHAR(20), APEPAT VARCHAR(20), APEMAT VARCHAR(20), FECHANAC DATETIME, EMPRESA VARCHAR(100), MES INT, DIA INT )"
    Set RS_GENERAL = New ADODB.Recordset
    X = 0
    RS_GENERAL.Open "SELECT * FROM EMPRESAS", DBSTARPLAN
    If RS_GENERAL.RecordCount Then
    While Not RS_GENERAL.EOF
            If xTodos.Value = 0 Then
                SQLINSERT = " INSERT INTO  [##_TMPCUMPLE" & VGL_COMPUTER & "]  (CODTRAB, APEPAT, APEMAT, NOMBRES, FECHANAC, MES, DIA, EMPRESA)  SELECT CODTRAB, APEPAT, APEMAT, NOMBRE, FECHANAC, MONTH(FECHANAC), DAY(FECHANAC), '" & RS_GENERAL.Fields(0) & "' FROM " & RS_GENERAL.Fields(2) & ".dbo.TRABAJADORES WHERE MONTH(FECHANAC)=" & xMes.Month & " ORDER BY FECHANAC "
            ElseIf xTodos.Value = 1 Then
                SQLINSERT = " INSERT INTO  [##_TMPCUMPLE" & VGL_COMPUTER & "]  (CODTRAB, NOMBRES, APEPAT, APEMAT, FECHANAC, MES, DIA, EMPRESA)  SELECT CODTRAB, APEPAT, APEMAT, NOMBRE, FECHANAC, MONTH(FECHANAC), DAY(FECHANAC), '" & RS_GENERAL.Fields(0) & "' FROM " & RS_GENERAL.Fields(2) & ".dbo.TRABAJADORES ORDER BY MONTH(FECHANAC) "
            End If
            DBSTARPLAN.Execute SQLINSERT
      RS_GENERAL.MoveNext
    Wend
  End If
  RS_GENERAL.Close 'ÓBTIENE LAS RUTAS DE LAS EMPRESAS
    DBSTARPLAN.Execute "EXECUTE [SP_CUMPLE] '" & VGL_COMPUTER & "'"
    'ABRE EL REPORTE
    With Reporte
       .Reset
        If Option1.Value Then
            .ReportFileName = REGSISTEMA.REPORTES & "\PLNAS007.RPT" ' 1 ARCHIVO DE REPORTE
            .WindowTitle = "PLNAS007 - CUMPLEAÑOS"
        Else
            .ReportFileName = REGSISTEMA.REPORTES & "\PLNAS008.RPT"  ' 2 ARCHIVO DE REPORTE
            .WindowTitle = "PLNAS008 - CUMPLEAÑOS"
        End If
        .Connect = "DSN=" & VGL_SERVER & ";User=sa;PWD=;DSQ=STARPLAN"
        .StoredProcParam(0) = VGL_COMPUTER
        .Formulas(1) = "XTITULO='" & IIf(xTodos.Value = 1, "Año " & Year(xAnno.Value), " Mes de " & AMESES(Month(xMes.Value)) & " de " & Year(xAnno.Value)) & "'"
            .SortFields(0) = "+{SP_CUMPLE.MES}"
            .SortFields(1) = "+{SP_CUMPLE.DIA}"
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .WindowShowPrintBtn = True
        .WindowShowRefreshBtn = True
        .WindowShowSearchBtn = True
        .WindowShowPrintSetupBtn = True
        If .Status <> 2 Then .Action = 1
    End With
End If
End Sub
Private Sub CMSELECTRAB_CLICK()
    REGSELECT.USARFECHACESE = False
    frSelect.Show 1
End Sub
Private Sub FORM_KEYDOWN(KEYCODE As Integer, Shift As Integer)
    If KEYCODE = vbKeyF5 Then CMSELECTRAB_CLICK
End Sub
Private Sub Form_Load()
    xAnno.Value = Date
    xMes.Value = Date
    Option4.Value = True
End Sub

Private Sub FORM_UNLOAD(CANCEL As Integer)
If ExisteTablaAux("_TMPTRCU") Then DBSYSTEM.Execute "DROP TABLE  _TMPTRCU"
End Sub

Private Sub OPTION3_Click()
If Option3.Value Then
  cmSelecTrab.Enabled = False
End If
End Sub

Private Sub OPTION4_Click()
If Option4.Value Then
  cmSelecTrab.Enabled = True
End If
End Sub
Private Sub XTODOS_Click()
    If xTodos.Value = 1 Then xMes.Visible = False Else xMes.Visible = True
End Sub

