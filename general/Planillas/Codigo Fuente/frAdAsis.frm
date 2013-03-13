VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C5FF36B7-A67E-11D3-9D9D-E6F193F7F854}#9.0#0"; "ApliCTxt.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frAdAsis 
   Caption         =   "Registro de Asistencia"
   ClientHeight    =   5670
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7005
   Icon            =   "frAdAsis.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5670
   ScaleMode       =   0  'User
   ScaleWidth      =   7005
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid DGAsis 
      Height          =   3540
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   6244
      _Version        =   393216
      AllowUpdate     =   -1  'True
      BackColor       =   16777215
      HeadLines       =   3
      RowHeight       =   18
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
      Caption         =   "Asistencia de:"
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
   Begin VB.Frame frameContenedorx1 
      BorderStyle     =   0  'None
      Height          =   825
      Left            =   0
      TabIndex        =   10
      Top             =   4875
      Width           =   6945
      Begin VB.CommandButton cmAceptar 
         Caption         =   "&Grabar"
         Height          =   345
         Left            =   4725
         TabIndex        =   15
         Top             =   240
         Width           =   930
      End
      Begin VB.CommandButton cmCancelar 
         Caption         =   "&Cerrar"
         Height          =   345
         Left            =   5730
         TabIndex        =   14
         Top             =   240
         Width           =   930
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Buscar"
         Height          =   345
         Left            =   3720
         TabIndex        =   13
         Top             =   240
         Width           =   930
      End
      Begin VB.CommandButton CmdImprimir 
         Caption         =   "&Imprimir"
         Height          =   345
         Left            =   2700
         TabIndex        =   12
         Top             =   240
         Width           =   930
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Hoja de Trabajo"
         Height          =   345
         Left            =   1005
         TabIndex        =   11
         Top             =   240
         Width           =   1530
      End
      Begin VB.Image Image1 
         Height          =   720
         Left            =   60
         Picture         =   "frAdAsis.frx":08CA
         Top             =   150
         Width           =   720
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   405
         Picture         =   "frAdAsis.frx":1384
         Top             =   285
         Width           =   480
      End
   End
   Begin Crystal.CrystalReport Reporte 
      Left            =   1095
      Top             =   5025
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame1 
      Caption         =   "&Trabajador"
      Height          =   1140
      Left            =   105
      TabIndex        =   0
      Top             =   75
      Width           =   6885
      Begin VB.CommandButton cmOriginalCC 
         Caption         =   "No"
         Height          =   300
         Left            =   6345
         TabIndex        =   9
         Top             =   705
         Width           =   375
      End
      Begin VB.CommandButton cmBuscaCC 
         Caption         =   "..."
         Height          =   300
         Left            =   5910
         TabIndex        =   8
         Top             =   705
         Width           =   375
      End
      Begin MSComCtl2.DTPicker xFecha 
         Height          =   285
         Left            =   1395
         TabIndex        =   7
         Top             =   345
         Visible         =   0   'False
         Width           =   3345
         _ExtentX        =   5900
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "dddd dd MMM yyyy"
         Format          =   61865987
         CurrentDate     =   36689
      End
      Begin AplisetControlText.Aplitext xTrab 
         Height          =   285
         Left            =   1395
         TabIndex        =   2
         Top             =   345
         Width           =   5325
         _ExtentX        =   9393
         _ExtentY        =   503
         Text            =   ""
      End
      Begin VB.Label LFecha 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         Height          =   195
         Left            =   135
         TabIndex        =   6
         Top             =   420
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Centro de Costo"
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   5
         Top             =   735
         Width           =   1140
      End
      Begin VB.Label xCCosto 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1395
         TabIndex        =   3
         Top             =   720
         Width           =   4455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nombre:"
         Height          =   195
         Left            =   135
         TabIndex        =   1
         Top             =   390
         Width           =   600
      End
   End
End
Attribute VB_Name = "frAdAsis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RSASIS As New ADODB.Recordset
Private Sub CMACEPTAR_CLICK()
    Dim X As Integer
    If RSASIS.RecordCount = 0 Then
        MsgBox "No existen registros por Grabar", vbCritical
        Exit Sub
    End If
    If frRegAsi.Op1(1).Value Then
        If xTrab.Tag <> "" Then
            MsgBox "Falta un Trabajador", vbInformation
            Exit Sub
        End If
    End If
    RSASIS.MoveFirst
    If frRegAsi.Op1(0) Then
        DBSYSTEM.Execute "DELETE FROM ASIS" & REGSISTEMA.ANNO & " WHERE DIA=" & DateSQL(xFecha.Value) & " AND CODTRAB IN (SELECT CODTRAB FROM  [##_TMPASIS" & VGL_COMPUTER & "] )"
    Else
        DBSYSTEM.Execute "DELETE FROM ASIS" & REGSISTEMA.ANNO & " WHERE CODTRAB='" & xTrab.Tag & "' AND DIA IN (SELECT FECHA FROM  [##_TMPASIS" & VGL_COMPUTER & "] )"
    End If
    Do While Not RSASIS.EOF
        If frRegAsi.Op1(0) Then 'SI ES POR TRABAJADOR
            For X = 2 To DGAsis.Columns.Count - 1
                If Not IsNull(RSASIS.Fields(DGAsis.Columns(X).DataField).Value) Then DBSYSTEM.Execute "INSERT INTO ASIS" & REGSISTEMA.ANNO & " (CODTRAB, DIA, CONCEPTO, VALOR, CCOSTO, ID_FECHAPAGO) VALUES ('" & RSASIS!CODTRAB & "'," & DateSQL(xFecha.Value) & ",'" & DGAsis.Columns(X).DataField & "'," & RSASIS.Fields(DGAsis.Columns(X).DataField).Value & ",' ',0)"
            Next
        Else
            For X = 1 To DGAsis.Columns.Count - 1
                If Not IsNull(RSASIS.Fields(DGAsis.Columns(X).DataField).Value) Then DBSYSTEM.Execute "INSERT INTO ASIS" & REGSISTEMA.ANNO & " (CODTRAB, DIA, CONCEPTO, VALOR, CCOSTO, ID_FECHAPAGO) VALUES ('" & xTrab.Tag & "'," & DateSQL(RSASIS!FECHA) & ",'" & DGAsis.Columns(X).DataField & "'," & RSASIS.Fields(DGAsis.Columns(X).DataField).Value & ",' ',0)"
            Next
        End If
        RSASIS.MoveNext
    Loop
    MsgBox "Información grabada Satisfactoriamente", vbInformation
End Sub
Private Sub CMBUSCACC_CLICK()
    Dim RSAUX As ADODB.Recordset
    Set RSAUX = New ADODB.Recordset
    RSAUX.Open "SELECT CODCCOSTO, NOMBRE FROM CCOSTOS ORDER BY CODCCOSTO", DBSYSTEM, adOpenStatic
    If RSAUX.RecordCount = 0 Then
        MsgBox "No se han encontrado regsitros de Centro de Costos", vbCritical
        Set RSAUX = Nothing
        Exit Sub
    End If
    frmComun.CONECTAR RSAUX
    frmComun.Show 1
    If VGUTIL(1) <> "" Then
        xCCosto.Caption = RSAUX!CODCCOSTO & " - " & RSAUX!NOMBRE
        xCCosto.Tag = RSAUX!CODCCOSTO
    End If
End Sub

Private Sub CMCANCELAR_CLICK()
    Unload Me
End Sub
Private Sub CmdImprimir_Click()
    'Imprimir Reporte
    Dim REG As Long
    
    
    RSASIS.Requery
    DGAsis.Refresh
        
    Screen.MousePointer = 11
    Call CREARPLAN(REG)
    DBSTARPLAN.Execute "EXECUTE [ASISTMP] ' [##TMPCREPLAN" & VGL_COMPUTER & "] '"
    With Reporte
        .Reset
        .ReportFileName = REGSISTEMA.REPORTES & "PLAN0064.RPT"
        .Connect = "DSN=" & VGL_SERVERREP & ";User=sa;PWD=;DSQ=" & VGL_BASE & ""
        .StoredProcParam(0) = "[##TMPCREPLAN" & VGL_COMPUTER & "] "
        .SortFields(0) = "+{ASISTMP.NOMBRES}"
        .WindowTitle = "PLAN0064.RPT -" & DGAsis.Caption
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .WindowShowPrintBtn = True
        .WindowShowRefreshBtn = True
        .WindowShowSearchBtn = True
        .WindowShowPrintSetupBtn = True
        .Formulas(0) = "XEMPRESA='" & REGSISTEMA.EMPRESA & "'"
        .Formulas(1) = "XRUC='" & REGSISTEMA.RUC & "'"
        .Formulas(2) = "XREG='" & Str(REG) & "'"
        If .Status <> 2 Then .Action = 1
    End With
    Screen.MousePointer = 1
End Sub
Private Sub CREARPLAN(Optional ByRef REG As Long)
    Dim RSTRABPLAN As New ADODB.Recordset
    Dim RSAUX As New ADODB.Recordset
    Dim I As Integer, ORDEN As Integer
    Dim CONC As String
    If ExisteTablaAux(" [##TMPCREPLAN" & VGL_COMPUTER & "] ") Then DBSYSTEM.Execute "DROP TABLE  [##TMPCREPLAN" & VGL_COMPUTER & "] "
    DBSYSTEM.Execute "CREATE TABLE  [##TMPCREPLAN" & VGL_COMPUTER & "]  (CODTRAB VARCHAR(8),NOMBRES varchar(50) ,CODCONCEP varchar(15),DESCONCEP varchar(40),ORDEN int,MONTO  Numeric(20,2) )"
    RSAUX.Open " [##_TMPASIS" & VGL_COMPUTER & "] ", DBSYSTEM, adOpenKeyset, adLockReadOnly
    REG = RSAUX.RecordCount
    RSTRABPLAN.Open " [##TMPCREPLAN" & VGL_COMPUTER & "] ", DBSYSTEM, adOpenKeyset, adLockOptimistic
    If RSAUX.RecordCount = 0 Then
        MsgBox "No existe ningún registro para imprimir la Planilla"
        Exit Sub
    End If
    RSAUX.MoveFirst
    Do While Not RSAUX.EOF
        For I = 2 To DGAsis.Columns.Count - 1
            ORDEN = ORDEN + 1
            If RSAUX.Fields(I).Value > 0 Then
                RSTRABPLAN.AddNew
                RSTRABPLAN!CODTRAB = RSAUX!CODTRAB
                RSTRABPLAN!NOMBRES = RSAUX!Trabajador
                RSTRABPLAN!CODCONCEP = Trim(RSAUX.Fields(I).Name)
                CONC = DevuelveValor("SELECT NOMBRE FROM CONCEPTOS WHERE CODIGO='" & _
                                     Trim(RSAUX.Fields(I).Name) & "'", DBSYSTEM)
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
Private Sub CREARPLAN2(Optional ByRef REG As Long)
Dim RSTRABPLAN As New ADODB.Recordset
    Dim RSAUX As New ADODB.Recordset
    Dim I As Integer, ORDEN As Integer
    Dim CONC As String
    If ExisteTablaAux(" [##TMPCREPLAN" & VGL_COMPUTER & "] ") Then DBSYSTEM.Execute "DROP TABLE  [##TMPCREPLAN" & VGL_COMPUTER & "] "
    DBSYSTEM.Execute "CREATE TABLE  [##TMPCREPLAN" & VGL_COMPUTER & "]  (CODTRAB VARCHAR(8),NOMBRES varchar(50) ,CODCONCEP varchar(15),DESCONCEP varchar(40),ORDEN int,MONTO  Numeric(20,2) )"
    RSAUX.Open " [##_TMPASIS" & VGL_COMPUTER & "] ", DBSYSTEM, adOpenKeyset, adLockReadOnly
    REG = RSAUX.RecordCount
    RSTRABPLAN.Open " [##TMPCREPLAN" & VGL_COMPUTER & "] ", DBSYSTEM, adOpenKeyset, adLockOptimistic
    
    If RSAUX.RecordCount = 0 Then
        MsgBox "No existe ningún registro para imprimir la Planilla"
        Exit Sub
    End If
    
    RSAUX.MoveFirst
    Do While Not RSAUX.EOF
        For I = 2 To DGAsis.Columns.Count - 1
            ORDEN = ORDEN + 1
            'If RSAUX.Fields(I).Value > 0 Then
                RSTRABPLAN.AddNew
                RSTRABPLAN!CODTRAB = RSAUX!CODTRAB
                RSTRABPLAN!NOMBRES = RSAUX!Trabajador
                RSTRABPLAN!CODCONCEP = Trim(RSAUX.Fields(I).Name)
                CONC = DevuelveValor("SELECT NOMBRE FROM CONCEPTOS WHERE CODIGO='" & _
                                     Trim(RSAUX.Fields(I).Name) & "'", DBSYSTEM)
                RSTRABPLAN!DESCONCEP = CONC
                RSTRABPLAN!ORDEN = ORDEN
                'RSTRABPLAN!MONTO = RSAUX.Fields(I).Value
                RSTRABPLAN.Update
            'End If
        Next
        ORDEN = 0
        RSAUX.MoveNext
    Loop
End Sub
Private Sub CMORIGINALCC_CLICK()
    xCCosto.Caption = ""
    xCCosto.Tag = ""
End Sub

Private Sub Command1_Click()
    If RSASIS.RecordCount = 0 Or RSASIS.EOF Then
        MsgBox "NO EXISTEN REGISTROS POR BUSCAR", vbInformation
        Exit Sub
    End If
    Dim RSAUX As New ADODB.Recordset
    Set RSAUX = RSASIS.Clone
    frmComun.CONECTAR RSAUX
    frmComun.Show 1
    If VGUTIL(1) <> "" Then
        RSASIS.MoveFirst
        RSASIS.FIND "CODTRAB='" & VGUTIL(1) & "'"
    End If
    Set RSAUX = Nothing
End Sub

Private Sub Command2_Click()
'Imprimir Reporte
    Dim REG As Long
    Screen.MousePointer = 11
    Call CREARPLAN2(REG)
    'DBSTARPLAN.Execute "EXECUTE [ASISTMP] '[##_TMPASIS" & VGL_COMPUTER & "]'"
    DBSTARPLAN.Execute "EXECUTE [ASISTMP] ' [##TMPCREPLAN" & VGL_COMPUTER & "] '"
    With Reporte
        .Reset
        .ReportFileName = REGSISTEMA.REPORTES & "PLAN0064.RPT"
        .Connect = "DSN=" & VGL_SERVERREP & ";User=sa;PWD=;DSQ=" & VGL_BASE & ""
        .StoredProcParam(0) = "[##TMPCREPLAN" & VGL_COMPUTER & "]"
        .SortFields(0) = "+{ASISTMP.NOMBRES}"
        .WindowTitle = "PLAN0064.RPT -" & DGAsis.Caption
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .WindowShowPrintBtn = True
        .WindowShowRefreshBtn = True
        .WindowShowSearchBtn = True
        .WindowShowPrintSetupBtn = True
        .Formulas(0) = "XEMPRESA='" & REGSISTEMA.EMPRESA & "'"
        .Formulas(1) = "XRUC='" & REGSISTEMA.RUC & "'"
        .Formulas(2) = "XREG='" & Str(REG) & "'"
        If .Status <> 2 Then .Action = 1
    End With
    Screen.MousePointer = 1
End Sub


Private Sub DGASIS_HEADCLICK(ByVal COLINDEX As Integer)
Static BIT_ORDEN As Boolean
    If COLINDEX > (IIf(frRegAsi.Op1(0), 1, 0)) Then
        frValor.Show 1
    Else
       If BIT_ORDEN = True Then
            RSASIS.Sort = DGAsis.Columns(COLINDEX).Caption & "  ASC"
            BIT_ORDEN = False
       Else
            RSASIS.Sort = DGAsis.Columns(COLINDEX).Caption & "  DESC"
            BIT_ORDEN = True
       End If
       Exit Sub
    End If
    If VPTAREA = "0" Then Exit Sub
    With RSASIS
            .MoveFirst
            Do While Not .EOF
                .Fields(DGAsis.Columns(COLINDEX).DataField).Value = Val(VPTAREA)
                .MoveNext
            Loop
            .MoveFirst
    End With
End Sub

Private Sub Form_Load()
    Dim RSAUX As New ADODB.Recordset
    RSAUX.Open "SELECT DISTINCT CONCEPTOS.* FROM CONCEPTOS WHERE CONCEPTOS.TIPO=0 AND CONCEPTOS.ESESCRITO=1 order by FILA ", DBSYSTEM, adOpenStatic
    Dim CAD As String
    If ExisteTablaAux(" [##_TMPASIS" & VGL_COMPUTER & "] ") Then DBSYSTEM.Execute "DROP TABLE  [##_TMPASIS" & VGL_COMPUTER & "] "
    If frRegAsi.Op1(0).Value Then 'SI ES POR TRABAJADOR
        CAD = "CREATE TABLE  [##_TMPASIS" & VGL_COMPUTER & "]  (CODTRAB VARCHAR(8), TRABAJADOR varchar(50)"
    Else                          'SI ES POR CONCEPTO
        CAD = "CREATE TABLE  [##_TMPASIS" & VGL_COMPUTER & "]  (FECHA DATETIME"
        Command1.Visible = False
    End If
    Do While Not RSAUX.EOF
        CAD = CAD & ", " & RSAUX!Codigo & "  Numeric(20,2) "
        RSAUX.MoveNext
    Loop
    CAD = CAD & ")"
    DBSYSTEM.Execute CAD
    RSASIS.Open " [##_TMPASIS" & VGL_COMPUTER & "] ", DBSYSTEM, adOpenKeyset, adLockOptimistic
    If frRegAsi.Op1(0).Value Then 'SI ES POR TRABAJADOR
        DBSYSTEM.Execute "INSERT INTO  [##_TMPASIS" & VGL_COMPUTER & "]  (CODTRAB, TRABAJADOR) SELECT CODTRAB, NOMBRES FROM  [##TMPSELECT" & VGL_COMPUTER & "]  ORDER BY NOMBRES"
        Lfecha.Visible = True
        xFecha.Visible = True
        xTrab.Visible = False
        Label1.Visible = False
        xFecha.Value = CDate(frRegAsi.xFechaIni)
        xFecha.MinDate = CDate(frRegAsi.xFechaIni)
        xFecha.MaxDate = CDate(frRegAsi.xFechaFin)
        XFECHA_CHANGE
    Else
        Dim NUMD As Integer, XFEC As Date, X As Integer
        NUMD = frRegAsi.xFechaFin.Value - frRegAsi.xFechaIni.Value
        XFEC = frRegAsi.xFechaIni.Value
        For X = 0 To NUMD
            DBSYSTEM.Execute "INSERT INTO  [##_TMPASIS" & VGL_COMPUTER & "]  (FECHA) VALUES (" & DateSQL(XFEC) & ")"
            XFEC = DateAdd("D", 1, XFEC)
        Next
    End If
    RSASIS.Requery
    Set DGAsis.DataSource = RSASIS
    If frRegAsi.Op1(0).Value Then
        DGAsis.Columns("CODTRAB").Locked = True
        DGAsis.Columns("TRABAJADOR").Locked = True
        REFRESCAR
    Else
        DGAsis.Columns("FECHA").Locked = True
    End If
    
    Call WHIT_DATAGRID(DGAsis)
End Sub

Private Sub Form_Resize()
If Me.Width < 7125 Then Exit Sub
If Me.Height < 6045 Then Exit Sub
'me.ScaleWidth=7005
'me.ScaleHeight=5640
'***********************************************
frameContenedorx1.TOP = Me.ScaleHeight - 765
'***********************************************
DGAsis.Width = Me.ScaleWidth - 150
DGAsis.Height = Me.ScaleHeight - 2100
'*********************************************
End Sub

Private Sub FORM_UNLOAD(CANCEL As Integer)
    Set RSASIS = Nothing
End Sub

Private Sub XFECHA_CHANGE()
    'On Error Resume Next
    Dim RSCARGA As New ADODB.Recordset
    If frRegAsi.Op1(0).Value Then
        RSCARGA.Open "SELECT CODTRAB, CONCEPTO, VALOR FROM ASIS2000 WHERE DIA=" & DateSQL(xFecha.Value) & " AND CODTRAB IN (SELECT CODTRAB FROM  [##TMPSELECT" & VGL_COMPUTER & "] )", DBSYSTEM, adOpenStatic
        Do While Not RSCARGA.EOF
            DBSYSTEM.Execute "UPDATE  [##_TMPASIS" & VGL_COMPUTER & "]  SET " & RSCARGA!CONCEPTO & "=" & RSCARGA!VALOR & " WHERE CODTRAB='" & RSCARGA!CODTRAB & "'"
            RSCARGA.MoveNext
        Loop
        Set RSCARGA = Nothing
        RSASIS.Requery
        Set DGAsis.DataSource = RSASIS
    End If
End Sub
Private Sub XTRAB_DBLCLICK()
    Dim RSTRAB As New ADODB.Recordset
    RSTRAB.Open "SELECT * FROM  [##TMPSELECT" & VGL_COMPUTER & "]  ORDER BY NOMBRES", DBSYSTEM, adOpenKeyset, adLockOptimistic
    frmComun.CONECTAR RSTRAB
    frmComun.Show 1
    If VGUTIL(1) <> "" Then
        xTrab.Text = RSTRAB!NOMBRES 'VGUTIL(1) & " :  " & VGUTIL(2)
        xCCosto.Caption = RSTRAB!AREA
        xCCosto.Tag = RSTRAB!CENTROCOSTO
        xTrab.Tag = VGUTIL(1)
    End If
    Set RSTRAB = Nothing
End Sub
Public Sub REFRESCAR()
    Dim X As Integer
    For X = 2 To DGAsis.Columns.Count - 1
        DGAsis.Columns(X).Alignment = dbgRight
        DGAsis.Columns(X).NumberFormat = "0.00 "
        DGAsis.Columns(X).Caption = "" & DevuelveValor("SELECT NOMBRE FROM CONCEPTOS WHERE CODIGO='" & DGAsis.Columns(X).Caption & "'", DBSYSTEM)
    Next
End Sub
Private Sub WHIT_DATAGRID(ByVal MEDATAGRID As DataGrid)
Dim K As Integer
    MEDATAGRID.Columns(0).Width = 750
    MEDATAGRID.Columns(1).Width = 2900
    For K = 2 To MEDATAGRID.Columns.Count - 1
        MEDATAGRID.Columns(K).Width = 800
    Next K
End Sub

