VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frCuentas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cuentas Corrientes"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7425
   Icon            =   "frCuentas.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   7425
   Tag             =   "Panel de Cuentas Corrientes"
   Begin VB.TextBox SqlCad 
      Height          =   345
      Left            =   4440
      TabIndex        =   16
      Text            =   "Text1"
      Top             =   135
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.CommandButton xCerrar 
      Cancel          =   -1  'True
      Caption         =   "&Cerrar"
      Height          =   375
      Left            =   5490
      TabIndex        =   14
      Top             =   120
      Width           =   1545
   End
   Begin VB.Frame Frame2 
      Caption         =   "Trabajadores con Saldos Pendientes"
      Height          =   2535
      Left            =   120
      TabIndex        =   8
      Top             =   3450
      Width           =   7155
      Begin VB.CommandButton cmAgregarProg 
         Caption         =   "Agregar Prog."
         Height          =   345
         Left            =   5340
         TabIndex        =   17
         Top             =   1950
         Width           =   1545
      End
      Begin Crystal.CrystalReport Reporte 
         Left            =   2910
         Top             =   1965
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.CommandButton cmEstado 
         Caption         =   "Historial"
         Enabled         =   0   'False
         Height          =   345
         Left            =   5340
         TabIndex        =   13
         Top             =   1506
         Width           =   1545
      End
      Begin VB.CommandButton cmTDel 
         Caption         =   "Eliminar"
         Height          =   345
         Left            =   5340
         TabIndex        =   12
         Top             =   1094
         Width           =   1545
      End
      Begin VB.CommandButton cmTEdit 
         Caption         =   "Editar"
         Height          =   345
         Left            =   5340
         TabIndex        =   11
         Top             =   682
         Width           =   1545
      End
      Begin VB.CommandButton cmTAdd 
         Caption         =   "Agregar"
         Height          =   345
         Left            =   5340
         TabIndex        =   10
         Top             =   270
         Width           =   1545
      End
      Begin MSDataGridLib.DataGrid dgSaldos 
         Height          =   2085
         Left            =   240
         TabIndex        =   9
         Top             =   300
         Width           =   4900
         _ExtentX        =   8652
         _ExtentY        =   3678
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
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
         ColumnCount     =   4
         BeginProperty Column00 
            DataField       =   "Nombres"
            Caption         =   "Nombre"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "Capital"
            Caption         =   "Capital"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "Saldo"
            Caption         =   "Saldo"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "PROGRAMADO"
            Caption         =   "PROG"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   3
            AllowRowSizing  =   0   'False
            BeginProperty Column00 
               ColumnWidth     =   1995.024
            EndProperty
            BeginProperty Column01 
               Alignment       =   1
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column02 
               Alignment       =   1
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column03 
               Alignment       =   2
               ColumnWidth     =   345.26
            EndProperty
         EndProperty
      End
   End
   Begin VB.ComboBox xTipo 
      Height          =   315
      ItemData        =   "frCuentas.frx":030A
      Left            =   2130
      List            =   "frCuentas.frx":0314
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   120
      Width           =   2145
   End
   Begin VB.Frame Frame1 
      Caption         =   "Grupo de Cuentas"
      Height          =   2835
      Left            =   120
      TabIndex        =   0
      Top             =   570
      Width           =   7155
      Begin VB.CommandButton cmEdit 
         Caption         =   "Editar"
         Height          =   345
         Left            =   5340
         TabIndex        =   15
         Top             =   761
         Width           =   1545
      End
      Begin VB.CommandButton cmCRpt 
         Caption         =   "&Reporte"
         Height          =   345
         Left            =   5340
         TabIndex        =   7
         Top             =   1683
         Width           =   1545
      End
      Begin VB.CommandButton cmCCon 
         Caption         =   "&Consolidado"
         Height          =   345
         Left            =   5340
         TabIndex        =   6
         Top             =   2145
         Width           =   1545
      End
      Begin VB.CommandButton cmCDel 
         Caption         =   "&Eliminar"
         Height          =   345
         Left            =   5340
         TabIndex        =   5
         Top             =   1230
         Width           =   1545
      End
      Begin VB.CommandButton cmCAdd 
         Caption         =   "&Agregar"
         Height          =   345
         Left            =   5340
         TabIndex        =   4
         Top             =   300
         Width           =   1545
      End
      Begin MSDataGridLib.DataGrid dgCuentas 
         Height          =   2385
         Left            =   240
         TabIndex        =   3
         Top             =   300
         Width           =   4900
         _ExtentX        =   8652
         _ExtentY        =   4207
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
               LCID            =   2058
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
               LCID            =   2058
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   3
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Tipo de Cuenta Corriente"
      Height          =   195
      Left            =   150
      TabIndex        =   1
      Top             =   180
      Width           =   1770
   End
End
Attribute VB_Name = "frCuentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents RSCUENTAS As ADODB.Recordset
Attribute RSCUENTAS.VB_VarHelpID = -1
Dim RSMOVIS As New ADODB.Recordset
Dim VTIPO As Byte
Dim ITSOPEN As Boolean
Dim ITSOPEN2 As Boolean
Dim REGACT As REGWIN

Private Sub CMAGREGARPROG_CLICK()
    If RSCUENTAS.EOF Then Exit Sub
    VPTAREA = RSCUENTAS!NOMBRE
    VPTRASPRM = RSCUENTAS!CODGRUPO
    frCuentasCtesProg.MANT = 0
    frCuentasCtesProg.Show 1
    RSMOVIS.Requery
End Sub
Private Sub CMCADD_CLICK()
    VPTAREA = "NUEVO"
    frECta.Show 1
    XTIPO_CLICK
End Sub
Public Sub IMPRIMIR()
    Dim X As Long
    Dim TIPO As String
    Screen.MousePointer = 11
    TIPO = ""
   If UCase(xTipo.Text) = UCase("INGRESOS") Then
        TIPO = " AND A.TIPOGRUPO=1 "
        Else:
            TIPO = " AND A.TIPOGRUPO=2 "
   End If
    If ExisteTablaAux(" [##TMPGRUPCTEPEND" & VGL_COMPUTER & "] ") Then DBSTARPLAN.Execute "DROP TABLE  [##TMPGRUPCTEPEND" & VGL_COMPUTER & "] "
    SqlCad.Text = "" & _
        "SELECT A.CODGRUPO,B.NOMBRE ," & _
        "SUM(A.CAPITAL) AS CAPITAL, SUM(A.SALDO)AS SALDO," & _
        "SUM(A.CUOTA) AS CUOTA " & _
        " INTO  [##TMPGRUPCTEPEND" & VGL_COMPUTER & "]  " & _
        " FROM [" & REGSISTEMA.BASESQL & "].dbo.MOVICTA A, [" & REGSISTEMA.BASESQL & "].dbo.CTAGRUPO B" & _
        " WHERE " & _
        "    A.CODGRUPO=B.CODGRUPO AND " & _
        "    A.SALDO<>0 " & TIPO & _
        " GROUP BY A.CODGRUPO,B.NOMBRE"
        
    DBSTARPLAN.Execute SqlCad.Text, X
    Screen.MousePointer = 1
    If X = 0 Then
      MsgBox "MENSAJE DEL SISTEMA: " & _
      " NO SE ENCONTRARÓN REGISTROS ", vbInformation
      Exit Sub
    End If
    DBSTARPLAN.Execute "EXECUTE [ASISTMP] ' [##TMPGRUPCTEPEND" & VGL_COMPUTER & "] '"
    With Reporte
        .Reset
        .ReportFileName = REGSISTEMA.REPORTES & "PLAN0015.RPT"
        .Connect = "DSN=" & VGL_SERVER & ";User=sa;PWD=;DSQ=" & REGSISTEMA.BASESQL & ""
        .StoredProcParam(0) = " [##TMPGRUPCTEPEND" & VGL_COMPUTER & "] "
        .Destination = crptToWindow
        .WindowShowPrintBtn = True
        .WindowShowRefreshBtn = True
        .WindowShowSearchBtn = True
        .WindowShowPrintSetupBtn = True
        .WindowTitle = "PLAN0015 - CONSOLIDADO DE SALDOS PENDIENTES POR GRUPO DE CTA."
        .Formulas(0) = "XEMPRESA='" & REGSISTEMA.EMPRESA & "'"
        .Formulas(1) = "XTIPO='" & xTipo.Text & "'"
        .Formulas(2) = "XRUC='" & REGSISTEMA.RUC & "'"
        .Formulas(3) = "XHORA='" & Format(Time, "HH:MM") & "'"
        If .Status <> 2 Then .Action = 1
    End With
End Sub

Private Sub CMCCON_CLICK()
    IMPRIMIR
End Sub

Private Sub cmCDel_Click()
    If DevuelveValor("SELECT DISTINCT CODGRUPO FROM MOVICTA WHERE CODGRUPO= '" & RSCUENTAS("CODGRUPO") & "' AND TIPOGRUPO=" & xTipo.ListIndex + 1, DBSYSTEM) = "" Then
        DBSYSTEM.Execute "DELETE FROM CTAGRUPO WHERE CODGRUPO='" & Trim(RSCUENTAS("CODGRUPO")) & "' AND TIPO=" & xTipo.ListIndex + 1
        RSCUENTAS.Requery
      Else:
        MsgBox "No se puede eliminar, esta siendo usado por algunos trabajadores en cta. cte", vbInformation
    End If
End Sub

Private Sub CMCRPT_CLICK()
    Dim X As Long
    Screen.MousePointer = 11
    If ExisteTablaAux(" [##TMPCTACTEPEND" & VGL_COMPUTER & "] ") Then DBSTARPLAN.Execute "DROP TABLE  [##TMPCTACTEPEND" & VGL_COMPUTER & "] "
    SqlCad.Text = "" & _
        "SELECT A.CODGRUPO,B.NOMBRE,A.CODMOV," & _
        "A.CODTRAB , " & _
        "C.APEPAT + ' '  + C.APEMAT + ' ' + C.NOMBRE AS NOMBRES, " & _
        "A.CAPITAL, " & _
        "A.INTERES, A.PORCQUINC, A.CUOTA, A.FECHAINI, A.NUMMESES, " & _
        "A.SALDO " & _
        " INTO  [##TMPCTACTEPEND" & VGL_COMPUTER & "]  " & _
        "FROM [" & REGSISTEMA.BASESQL & "].dbo.MOVICTA A,[" & REGSISTEMA.BASESQL & "].dbo.CTAGRUPO B,[" & REGSISTEMA.BASESQL & "].dbo.TRABAJADORES C " & _
        "WHERE A.CODGRUPO=B.CODGRUPO AND " & _
        " A.CODTRAB = C.CODTRAB AND A.CODGRUPO='" & _
        RSCUENTAS!CODGRUPO & "' AND A.SALDO<>0"
    DBSTARPLAN.Execute SqlCad.Text, X
    Screen.MousePointer = 1
    If X = 0 Then
      MsgBox "MENSAJE DEL SISTEMA: " & _
      " NO SE ENCONTRARÓN REGISTROS ", vbInformation
      Exit Sub
    End If
     DBSTARPLAN.Execute "EXECUTE [ASISTMP] ' [##TMPCTACTEPEND" & VGL_COMPUTER & "] '"
    With Reporte
        .Reset
        .ReportFileName = REGSISTEMA.REPORTES & "PLAN0028.RPT"
        .Connect = "DSN=" & VGL_SERVERREP & ";User=sa;PWD=;DSQ=" & VGL_BASE & ""
        .StoredProcParam(0) = " [##TMPCTACTEPEND" & VGL_COMPUTER & "] "
        .Destination = crptToWindow
        .WindowShowPrintBtn = True
        .WindowShowRefreshBtn = True
        .WindowShowSearchBtn = True
        .WindowShowPrintSetupBtn = True
        .WindowTitle = "PLAN0028 - TRABAJADORES CON SALDOS PENDIENTES "
        .Formulas(0) = "XEMPRESA='" & REGSISTEMA.EMPRESA & "'"
        .Formulas(1) = "XHORA='" & Format(Time, "HH:MM") & "'"
        .Formulas(2) = "XTIPO='" & xTipo.Text & "'"
        .Formulas(3) = "XRUC='" & REGSISTEMA.RUC & "'"
        If .Status <> 2 Then .Action = 1
    End With
End Sub

Private Sub CMEDIT_CLICK()
    If RSCUENTAS.EOF Then Exit Sub
    VPTAREA = "EDITAR"
    Load frECta
    With frECta
        .xCodigo.Text = RSCUENTAS!CODGRUPO
        .xNombre.Text = RSCUENTAS!NOMBRE
        .xPlanilla.Tag = RSCUENTAS!PLANILLA
    End With
    frECta.Show 1
    XTIPO_CLICK
End Sub

Private Sub CMESTADO_CLICK()
    Dim X As Long
    Screen.MousePointer = 11
    If ExisteTablaAux(" [##TMPCTEHIST" & VGL_COMPUTER & "] ") Then DBAUXCOM.Execute "DROP TABLE  [##TMPCTEHIST" & VGL_COMPUTER & "] "
    SqlCad.Text = "" & _
        " SELECT PAGOSCTA.*,NOMBOL.NOMBRE, NOMBOL.MES, " & _
        " NOMBOL.FECHAINI , NOMBOL.FECHAFIN, NOMBOL.FECHAPAGO, " & _
        " NOMBOL.DARADELANTO , NOMBOL.FECHAADELANTO, NOMBOL.CERRADO " & _
        " INTO   [##TMPCTEHIST" & VGL_COMPUTER & "]  " & _
        " FROM PAGOSCTA, NOMBOL " & _
        " WHERE PAGOSCTA.CODNOMBOL = [NOMBOL].[CODIGO]  AND " & _
        " PAGOSCTA.CODTRAB='" & RSMOVIS!CODTRAB & "'" & _
        " ORDER BY NOMBOL.FECHAPAGO "
    DBSYSTEM.Execute SqlCad.Text, X
    If X = 0 Then
      MsgBox "MENSAJE DEL SISTEMA: " & _
      " NO SE ENCONTRARÓN REGISTROS ", vbInformation
      Screen.MousePointer = 1
      Exit Sub
    End If
    DBSTARPLAN.Execute "EXECUTE [ASISTMP] ' [##TMPCTEHIST" & VGL_COMPUTER & "] '"
    With Reporte
        .Reset
        .ReportFileName = REGSISTEMA.REPORTES & "\PLAN0016.RPT"
        .Connect = "DSN=" & VGL_SERVER & ";User=sa;PWD=;DSQ=" & REGSISTEMA.BASESQL & ""
        .StoredProcParam(0) = " [##TMPCTEHIST" & VGL_COMPUTER & "] "
        .Destination = crptToWindow
        .WindowShowPrintBtn = True
        .WindowShowRefreshBtn = True
        .WindowShowSearchBtn = True
        .WindowShowPrintSetupBtn = True
        .WindowTitle = "PLAN0016 - HISTORIAL DE PAGOS A CUENTA"
        .Formulas(0) = "XEMPRESA='" & REGSISTEMA.EMPRESA & "'"
        .Formulas(1) = "XNOMEMP='" & RSMOVIS!NOMBRES & "'"
        .Formulas(2) = "XRUC='" & REGSISTEMA.RUC & "'"
        .Formulas(3) = "XHORA='" & Format(Time, "HH:MM") & "'"
        If .Status <> 2 Then .Action = 1
    End With
    Screen.MousePointer = 1
End Sub

Private Sub CMTADD_CLICK()
    If RSCUENTAS.EOF Then Exit Sub
    VPTAREA = "NUEVO"
    VPTRASPRM = RSCUENTAS!CODGRUPO
    frMoviCta.Show 1
    RSMOVIS.Requery
End Sub

Private Sub CMTDEL_CLICK()
    Dim RSAUX As New ADODB.Recordset
    RSAUX.Open "SELECT * FROM PAGOSCTA WHERE CODMOV=" & RSMOVIS!CODMOV, DBSYSTEM, adOpenKeyset
    If RSAUX.RecordCount > 0 Then
        MsgBox "Esta cuenta no se puede eliminar pues se han efectuado movimientos en adelantos o planilla", vbCritical
    Else
        If MsgBox("Realmente desea eliminar el registro seleccionado", vbYesNo + vbQuestion) = vbYes Then
            DBSYSTEM.Execute "DELETE FROM MOVICTA WHERE CODMOV=" & RSMOVIS!CODMOV & " AND TIPOGRUPO=" & xTipo.ListIndex + 1
            DBSYSTEM.Execute "DELETE FROM CTACTEPROG WHERE CODMOV=" & RSMOVIS!CODMOV
            RSMOVIS.Requery
        End If
    End If
    Set RSAUX = Nothing
End Sub

Private Sub CMTEDIT_CLICK()
    If RSCUENTAS.EOF Or RSCUENTAS.BOF Then Exit Sub
    If RSMOVIS.EOF Then Exit Sub
    'VERIFICAR QUE SE PUEDE EDITAR
    Dim RSAUX As New ADODB.Recordset
    RSAUX.Open "SELECT * FROM PAGOSCTA WHERE CODMOV=" & RSMOVIS!CODMOV, DBSYSTEM, adOpenKeyset
    
    If RSMOVIS!PROGRAMADO = 1 Then
        VPTAREA = RSCUENTAS!NOMBRE
        VPTRASPRM = RSCUENTAS!CODGRUPO
        VPCODTMP = RSMOVIS!CODMOV
        VPNUMTMP = RSMOVIS!CODTRAB
'        If RSAUX.RecordCount > 0 Then
'            MsgBox "Esta cuenta no se puede modificar pues se han efectuado movimientos en adelantos o planilla" & Chr(13) & "Solo podra Examinar ", vbInformation
'            frCuentasCtesProg.MANT = 2
'          Else
'            frCuentasCtesProg.MANT = 1
'        End If
        frCuentasCtesProg.MANT = 1
        frCuentasCtesProg.Frame2.Enabled = True
        frCuentasCtesProg.xTrab.Text = RSMOVIS!NOMBRES
        frCuentasCtesProg.Show 1
        RSMOVIS.Requery
    Else
        VPTAREA = "EDITAR"
        VPTRASPRM = RSMOVIS!NOMBRES
        VPCODTMP = RSMOVIS!CODMOV
        frMoviCta.Show 1
        RSMOVIS.Requery
    End If
End Sub

Private Sub Form_Activate()
    ActivarTools REGACT
End Sub

Private Sub Form_Load()
    Set RSCUENTAS = New ADODB.Recordset
    ITSOPEN = False
    ITSOPEN2 = False
    xTipo.ListIndex = 0
    With REGACT
        .BUSCAR = False
        .EDITAR = True
        .ELIMINAR = True
        .FILTRAR = False
        .IMPRIMIR = True
        .NUEVO = True
        .PRELIMINAR = True
    End With
End Sub

Private Sub FORM_UNLOAD(CANCEL As Integer)
    Set RSCUENTAS = Nothing
    Set RSMOVIS = Nothing
End Sub

Private Sub RSCUENTAS_MOVECOMPLETE(ByVal ADREASON As ADODB.EventReasonEnum, ByVal PERROR As ADODB.Error, ADSTATUS As ADODB.EventStatusEnum, ByVal PRECORDSET As ADODB.Recordset)
    If RSCUENTAS.EOF Or RSCUENTAS.BOF Then Exit Sub
    If ITSOPEN2 Then RSMOVIS.Close
    ITSOPEN2 = True
    Dim STRSQL As String
    STRSQL = "SELECT MOVICTA.CODMOV,MOVICTA.CODTRAB, TRABAJADORES.APEPAT + ' ' + TRABAJADORES.APEMAT + ' ' + TRABAJADORES.NOMBRE AS NOMBRES, MOVICTA.CAPITAL, MOVICTA.SALDO, MOVICTA.PROGRAMADO FROM TRABAJADORES INNER JOIN MOVICTA ON TRABAJADORES.CODTRAB = MOVICTA.CODTRAB WHERE SALDO<>0 AND CODGRUPO='" & RSCUENTAS!CODGRUPO & "' AND TIPOGRUPO=" & xTipo.ListIndex + 1
    'STRSQL = "SELECT MOVICTA.CODMOV,MOVICTA.CODTRAB, TRIM([TRABAJADORES]![APEPAT]) & ' ' & TRIM([TRABAJADORES]![APEMAT]) & ' ' & TRIM([TRABAJADORES]![NOMBRE]) AS NOMBRES, MOVICTA.CAPITAL, MOVICTA.SALDO FROM TRABAJADORES INNER JOIN MOVICTA ON TRABAJADORES.CODTRAB = MOVICTA.CODTRAB WHERE SALDO<>0 AND CODGRUPO='" & RSCUENTAS!CODGRUPO & "'"
    If RSMOVIS.State = 1 Then RSMOVIS.Close
    RSMOVIS.Open STRSQL, DBSYSTEM, adOpenKeyset
    Set dgSaldos.DataSource = RSMOVIS
    If RSMOVIS.RecordCount > 0 Then
        RSMOVIS.MoveFirst
        cmEstado.Enabled = True
      Else: cmEstado.Enabled = False
    End If
End Sub


Private Sub XCERRAR_CLICK()
    Unload Me
End Sub

Private Sub XTIPO_CLICK()
    Set RSCUENTAS = Nothing
    Set RSCUENTAS = New ADODB.Recordset
    RSCUENTAS.Open "SELECT * FROM CTAGRUPO WHERE TIPO=" & (xTipo.ListIndex + 1) & " ORDER BY NOMBRE", DBSYSTEM, adOpenKeyset, adLockPessimistic
    ITSOPEN = True
    Set dgCuentas.DataSource = RSCUENTAS
    dgCuentas.Columns("TIPO").Visible = False
    Dim STRSQL As String
    If RSCUENTAS.RecordCount = 0 Then
        STRSQL = "SELECT MOVICTA.CODMOV,MOVICTA.CODTRAB, TRABAJADORES.APEPAT + ' ' +TRABAJADORES.APEMAT + ' ' + TRABAJADORES.NOMBRE AS NOMBRES, MOVICTA.CAPITAL, MOVICTA.SALDO,MOVICTA.PROGRAMADO FROM TRABAJADORES INNER JOIN MOVICTA ON TRABAJADORES.CODTRAB = MOVICTA.CODTRAB WHERE SALDO<>0 AND CODGRUPO='00'"
        If RSMOVIS.State = 1 Then RSMOVIS.Close
        RSMOVIS.Open STRSQL, DBSYSTEM, adOpenKeyset, adLockOptimistic
        Set dgSaldos.DataSource = RSMOVIS
    End If
End Sub
Public Function EXISTE(ByVal STRCODGRP As String) As Boolean
    If RSCUENTAS.RecordCount = 0 Then
        EXISTE = False
        Exit Function
    End If
    RSCUENTAS.FIND "CODGRUPO='" & STRCODGRP & "'"
    If RSCUENTAS.EOF Then
        EXISTE = False
    Else
        EXISTE = True
    End If
End Function




