VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frInputMov 
   Caption         =   "Ingreso de Datos"
   ClientHeight    =   5865
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9840
   Icon            =   "frInputMov.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5865
   ScaleWidth      =   9840
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frameContenedorx1 
      BorderStyle     =   0  'None
      Height          =   810
      Left            =   0
      TabIndex        =   1
      Top             =   5130
      Width           =   9930
      Begin VB.CommandButton cmGrabar 
         Caption         =   "&Grabar"
         Height          =   375
         Left            =   5700
         TabIndex        =   6
         Top             =   210
         Width           =   1200
      End
      Begin VB.CommandButton cmSalir 
         Caption         =   "&Salir"
         Height          =   375
         Left            =   7020
         TabIndex        =   5
         ToolTipText     =   "Sale sn grabar los datos"
         Top             =   210
         Width           =   1200
      End
      Begin VB.CommandButton cmTotales 
         Caption         =   "&Totales"
         Height          =   375
         Left            =   90
         TabIndex        =   4
         Top             =   210
         Width           =   1200
      End
      Begin VB.CommandButton cmBuscar 
         Caption         =   "&Buscar"
         Height          =   375
         Left            =   4410
         TabIndex        =   3
         Top             =   210
         Width           =   1200
      End
      Begin VB.CommandButton CmdImprimir 
         Cancel          =   -1  'True
         Caption         =   "&Imprimir"
         Height          =   375
         Left            =   8550
         TabIndex        =   2
         ToolTipText     =   "Sale sn grabar los datos"
         Top             =   210
         Width           =   1200
      End
      Begin VB.Label xVaca 
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   1350
         TabIndex        =   7
         Top             =   225
         Visible         =   0   'False
         Width           =   3015
      End
   End
   Begin Crystal.CrystalReport Reporte 
      Left            =   2130
      Top             =   2880
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSDataGridLib.DataGrid DGLista 
      Height          =   4950
      Left            =   105
      TabIndex        =   0
      Top             =   135
      Width           =   9675
      _ExtentX        =   17066
      _ExtentY        =   8731
      _Version        =   393216
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
End
Attribute VB_Name = "frInputMov"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RSINPUT As New ADODB.Recordset

Private Sub CMBUSCAR_Click()
    If RSINPUT.RecordCount = 0 Or RSINPUT.EOF Then
        MsgBox "NO EXISTEN REGISTROS POR BUSCAR", vbInformation
        Exit Sub
    End If
    Dim RSAUX As New ADODB.Recordset
    Set RSAUX = RSINPUT.Clone
    frmComun.CONECTAR RSAUX
    frmComun.Show 1
    If VGUTIL(1) <> "" Then
        RSINPUT.MoveFirst
        RSINPUT.FIND "CODTRAB='" & VGUTIL(1) & "'"
    End If
    Set RSAUX = Nothing
    dgLista.SetFocus
End Sub

Private Sub CmdImprimir_Click()
    Dim REG As Long
    Screen.MousePointer = 11
    Call CREARPLAN(REG)
'    DBSTARPLAN.Execute "EXECUTE [SP_REMUNERA] 'TMPCREPLAN" & VGL_COMPUTER & "'"
    With Reporte
        .Reset
        .WindowTitle = "PLAN0065.RPT -" & dgLista.Caption
        .ReportFileName = REGSISTEMA.REPORTES & "PLAN0065.RPT"
        .Connect = "DSN=" & VGL_SERVERREP & ";User=sa;PWD=;DSQ=" & VGL_BASE & ""
        .StoredProcParam(0) = "##TMPCREPLAN" & VGL_COMPUTER & ""
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
    If ExisteTablaSQL("TMPCREPLAN", DBSTARPLAN) Then DBSTARPLAN.Execute "DROP TABLE TMPCREPLAN"
End Sub
Private Sub CREARPLAN(Optional ByRef REG As Long)
    Dim RSTRABPLAN As New ADODB.Recordset
    Dim RSAUX As New ADODB.Recordset
    Dim I As Integer, ORDEN As Integer
    Dim CONC As String
    If ExisteTablaSQL("##TMPCREPLAN" & VGL_COMPUTER & "", DBSYSTEM) Then DBSTARPLAN.Execute "DROP TABLE  [##TMPCREPLAN" & VGL_COMPUTER & "] "
    DBSTARPLAN.Execute "CREATE TABLE  ##TMPCREPLAN" & VGL_COMPUTER & "  (CODTRAB VARCHAR(8),NOMBRES VARCHAR(50) ,CODCONCEP VARCHAR(15),DESCONCEP VARCHAR(40),ORDEN INT,MONTO  Numeric(20,2) )"
    RSAUX.Open "##INPUTMOV" & VGL_COMPUTER & "", DBSTARPLAN, adOpenKeyset, adLockReadOnly
    REG = RSAUX.RecordCount
    RSTRABPLAN.Open "##TMPCREPLAN" & VGL_COMPUTER & "", DBSTARPLAN, adOpenKeyset, adLockOptimistic
    If RSAUX.RecordCount = 0 Then
        MsgBox "NO EXISTE NINGÚN REGISTRO PARA IMPRIMIR LA PLANILLA"
        Exit Sub
    End If
    RSAUX.MoveFirst
    Do While Not RSAUX.EOF
        For I = 2 To RSAUX.Fields.Count - 1
            ORDEN = ORDEN + 1
            If RSAUX.Fields(I).Value > 0 Then
                RSTRABPLAN.AddNew
                RSTRABPLAN!CODTRAB = RSAUX!CODTRAB
                RSTRABPLAN!NOMBRES = RSAUX!NOMBRES
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


Private Sub CMGRABAR_CLICK()
    Dim X As Integer
    If RSINPUT.RecordCount = 0 Then
        MsgBox "NO EXISTE NADA POR GRABAR", vbInformation
        Exit Sub
    End If
    RSINPUT.MoveFirst
    If MsgBox("CONFIRMA QUE DESEA GRABAR LOS DATOS EN LA BASE DE DATOS PRINCIPAL DEL SISTEMA", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    DBSYSTEM.Execute "DELETE FROM INGMOV2000 WHERE CODNOMBOL=" & REGINGMOV.CODNOMBOL & " AND CODTRAB IN (SELECT CODTRAB FROM  [##INPUTMOV" & VGL_COMPUTER & "] )"
    Do While Not RSINPUT.EOF
        For X = 2 To dgLista.Columns.Count - 1
            If Not IsNull(RSINPUT.Fields(dgLista.Columns(X).DataField).Value) Then DBSYSTEM.Execute "INSERT INTO INGMOV2000 (CODTRAB, CONCEPTO, VALOR, CODNOMBOL) VALUES ('" & RSINPUT!CODTRAB & "','" & dgLista.Columns(X).DataField & "'," & RSINPUT.Fields(dgLista.Columns(X).DataField).Value & "," & REGINGMOV.CODNOMBOL & ")"
        Next
        RSINPUT.MoveNext
    Loop
    MsgBox "LOS DATOS SE GRABARON SATISFACTORIAMENTE, SE PROCEDERÁ A CERRAR LA VENTANA", vbInformation
    Unload Me
End Sub

Private Sub CMSALIR_CLICK()
    Unload Me
End Sub

Private Sub CMTOTALES_Click()
    If RSINPUT.EOF Then Exit Sub
    VPTAREA = "FRINPUTMOV"
    RSINPUT.MoveFirst
    frSuma.Show 1
End Sub

Private Sub DGLISTA_HEADCLICK(ByVal COLINDEX As Integer)
    If COLINDEX > 1 Then
        frValor.Show 1
    Else
        RSINPUT.Sort = dgLista.Columns(COLINDEX).DataField
        Exit Sub
    End If
    If VPTAREA <> "0" Then
        With RSINPUT
                .MoveFirst
                Do While Not .EOF
                    .Fields(Trim$(dgLista.Columns(COLINDEX).DataField)).Value = Val(VPTAREA)
                    .MoveNext
                Loop
                .MoveFirst
        End With
    End If
End Sub

Private Sub Form_Load()
    Dim RSCNPT As New ADODB.Recordset, STRCREA As String
   RSCNPT.Open "SELECT * FROM CONCEPTOS WHERE ESESCRITO=1 AND TIPO<>0 AND (NOT CODIGO LIKE 'XX%') ORDER BY CODIGO", DBSYSTEM, adOpenStatic
    If ExisteTablaAux(" [##INPUTMOV" & VGL_COMPUTER & "] ") Then DBSTARPLAN.Execute "DROP TABLE  [##INPUTMOV" & VGL_COMPUTER & "] "
    STRCREA = "CREATE TABLE  [##INPUTMOV" & VGL_COMPUTER & "]  (CODTRAB VARCHAR(8), NOMBRES VARCHAR(100)"
    With RSCNPT
        .MoveFirst
        Do While Not .EOF
            STRCREA = STRCREA & ", " & RSCNPT!Codigo & "  Numeric(20,2) "
            .MoveNext
        Loop
    End With
    STRCREA = STRCREA & ")"
    DBSTARPLAN.Execute STRCREA
    Set RSCNPT = Nothing
    'SET RSCNPT = NEW ADODB.RECORDSET
    DBSTARPLAN.Execute "INSERT INTO  [##INPUTMOV" & VGL_COMPUTER & "]  (CODTRAB, NOMBRES) SELECT CODTRAB, NOMBRES FROM  [##TMPSELECT" & VGL_COMPUTER & "]  " & REGINGMOV.CADCONDI & " ORDER BY NOMBRES"
    'RECICLAJE DE RSCNPT
    If MsgBox("DESEA CARGAR LOS DATOS INGRESADOS ANTERIORMENTE", vbQuestion + vbYesNo) = vbYes Then
        RSCNPT.Open "SELECT * FROM INGMOV2000 WHERE CODNOMBOL=" & REGINGMOV.CODNOMBOL & " AND CODTRAB IN (SELECT CODTRAB FROM  [##INPUTMOV" & VGL_COMPUTER & "] )", DBSYSTEM, adOpenStatic
        Do While Not RSCNPT.EOF
            DBSTARPLAN.Execute "UPDATE  [##INPUTMOV" & VGL_COMPUTER & "]  SET " & RSCNPT!CONCEPTO & "=" & RSCNPT!VALOR & " WHERE CODTRAB='" & RSCNPT!CODTRAB & "'"
            RSCNPT.MoveNext
        Loop
        Set RSCNPT = Nothing
    End If
    'IMPORTANTE: TENEMOS QUE RECICLAR LA MAYOR CANTIDAD DE RECORDSETS
    'JALAMOS LAS VACACIONES
    RSCNPT.Open "SELECT * FROM HISTOVAC WHERE NOMBOL=" & REGINGMOV.CODNOMBOL & " AND CODTRAB IN (SELECT CODTRAB FROM  [##INPUTMOV" & VGL_COMPUTER & "] )", DBSYSTEM, adOpenStatic, adLockReadOnly
    If RSCNPT.RecordCount > 0 Or Not RSCNPT.EOF Then
        Do While Not RSCNPT.EOF
            DBSTARPLAN.Execute "UPDATE  [##INPUTMOV" & VGL_COMPUTER & "]  SET REMUVAC=" & RSCNPT!MONTO & " WHERE CODTRAB='" & RSCNPT!CODTRAB & "'"
            RSCNPT.MoveNext
        Loop
        xVaca.Visible = True
        xVaca.Caption = "** " & RSCNPT.RecordCount & " VACACIONES"
    End If
    Set RSCNPT = Nothing
    'JALAMOS LAS GRATIFICACIONES, SI LAS HAY PARA ESTE PERIODO DE PAGO
    RSCNPT.Open "SELECT PLANGRATI.* FROM PLANGRATI, GRATIFICACION WHERE GRATIFICACION.PERIODO=" & REGINGMOV.CODNOMBOL & " AND PLANGRATI.CODTRAB IN (SELECT CODTRAB FROM  [##INPUTMOV" & VGL_COMPUTER & "] )", DBSYSTEM, adOpenStatic, adLockReadOnly
    If RSCNPT.RecordCount > 0 Or Not RSCNPT.EOF Then
        Do While Not RSCNPT.EOF
            DBSTARPLAN.Execute "UPDATE  [##INPUTMOV" & VGL_COMPUTER & "]  SET REMUGRAT=" & RSCNPT!IMPORTEGRATI & " WHERE CODTRAB='" & RSCNPT!CODTRAB & "'"
            RSCNPT.MoveNext
        Loop
        xVaca.Visible = True
        xVaca.Caption = "** " & RSCNPT.RecordCount & " GRATIFICACIONES"
    End If
    Set RSCNPT = Nothing
    
    RSINPUT.Open " [##INPUTMOV" & VGL_COMPUTER & "] ", DBSTARPLAN, adOpenDynamic, adLockOptimistic
    Set dgLista.DataSource = RSINPUT
    dgLista.Columns("CODTRAB").Locked = True
    dgLista.Columns("NOMBRES").Locked = True
    REFRESCAR
End Sub

Private Sub Form_Resize()

If Me.Width < 9960 Then Exit Sub
If Me.Height < 6270 Then Exit Sub
'me.ScaleHeight=5865
'me.ScaleWidth=9840
'************************************
frameContenedorx1.TOP = Me.ScaleHeight - 735
'************************************
dgLista.Width = Me.ScaleWidth - 165
dgLista.Height = Me.ScaleHeight - 915

End Sub

Private Sub FORM_UNLOAD(CANCEL As Integer)
    Set RSINPUT = Nothing
End Sub

Public Sub REFRESCAR()
On Error Resume Next
    Dim X As Integer
    dgLista.Columns("REMUVAC").Locked = True
    For X = 2 To dgLista.Columns.Count - 1
        dgLista.Columns(X).Alignment = dbgRight
        dgLista.Columns(X).NumberFormat = "0.00 "
        dgLista.Columns(X).Caption = "" & DevuelveValor("SELECT NOMBRE FROM CONCEPTOS WHERE CODIGO='" & dgLista.Columns(X).Caption & "'", DBSYSTEM)
        dgLista.Columns(X).Width = 950
    Next
End Sub


