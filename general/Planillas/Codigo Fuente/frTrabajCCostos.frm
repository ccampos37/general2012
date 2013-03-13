VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frTrabajCCostos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Trabajadores en Varios Centros de Costos"
   ClientHeight    =   4305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9000
   Icon            =   "frTrabajCCostos.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   9000
   Begin Crystal.CrystalReport Reporte 
      Left            =   3750
      Top             =   1770
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmImprimir 
      Caption         =   "&Imprimir Listado"
      Height          =   480
      Left            =   1755
      TabIndex        =   6
      Top             =   3735
      Width           =   1455
   End
   Begin VB.CommandButton cmAddTrab 
      Caption         =   "Agregar &Trabajador"
      Height          =   480
      Left            =   90
      TabIndex        =   5
      Top             =   3735
      Width           =   1455
   End
   Begin VB.CommandButton cmCerrar 
      Cancel          =   -1  'True
      Caption         =   "&Cerrar"
      Height          =   360
      Left            =   7515
      TabIndex        =   4
      Top             =   3690
      Width           =   1305
   End
   Begin VB.CommandButton cmQuitar 
      Caption         =   "&Quitar"
      Height          =   360
      Left            =   6075
      TabIndex        =   3
      Top             =   3690
      Width           =   1305
   End
   Begin VB.CommandButton cmAgregar 
      Caption         =   "&Agregar"
      Height          =   360
      Left            =   4590
      TabIndex        =   2
      Top             =   3690
      Width           =   1305
   End
   Begin MSDataGridLib.DataGrid xDetalles 
      Height          =   3495
      Left            =   4590
      TabIndex        =   1
      Top             =   135
      Width           =   4230
      _ExtentX        =   7461
      _ExtentY        =   6165
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   2
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
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "CodCCosto"
         Caption         =   "Codigo Centro de Costo"
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
         DataField       =   "Nombre"
         Caption         =   "Nombre Descriptivo"
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
      BeginProperty Column02 
         DataField       =   "Basico"
         Caption         =   "Básico"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0.00 "
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
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2399.811
         EndProperty
         BeginProperty Column02 
            Alignment       =   1
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid xTrabs 
      Height          =   3495
      Left            =   90
      TabIndex        =   0
      Top             =   135
      Width           =   4515
      _ExtentX        =   7964
      _ExtentY        =   6165
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   2
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
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "CodTrab"
         Caption         =   "Codigo"
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
         DataField       =   "NomTrab"
         Caption         =   "Trabajadores Afectos"
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
            ColumnWidth     =   794.835
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   3105.071
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frTrabajCCostos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents RSTRABS As ADODB.Recordset
Attribute RSTRABS.VB_VarHelpID = -1
Dim RSDET As New ADODB.Recordset
Private Sub CMADDTRAB_CLICK()
    Set RSDET = Nothing
    RSDET.Open "SELECT CODTRAB, NOMBRES FROM VWTRABAJ WHERE CODTRAB NOT IN (SELECT DISTINCT CODTRAB FROM TRABXCOSTO)", DBSYSTEM, adOpenStatic, adLockReadOnly
    If RSDET.EOF Then
        MsgBox "No existen Trabajadores ó todos se encuentran ya registrados en varios Centros de Costos", vbExclamation
        Set RSDET = Nothing
        If RSTRABS.RecordCount > 0 Then RSTRABS.MOVE 0
        Exit Sub
    End If
    Dim OLDCOD As String
    frmComun.CONECTAR RSDET
    frmComun.Show 1
    If VGUTIL(1) <> "" Then
        OLDCOD = VGUTIL(1)
        Set RSDET = Nothing
        'RSDET.OPEN "SELECT * FROM CCOSTOS WHERE CODCCOSTO<>'" & DEVUELVEVALOR("SELECT CCOSTO FROM TRABAJADORES WHERE CODTRAB='" & OLDCOD & "'", DBSYSTEM) & "' AND CODCCOSTO NOT IN (SELECT CODCCOSTO FROM TRABXCOSTO WHERE CODTRAB='" & OLDCOD & "')", DBSYSTEM, ADOPENSTATIC, ADLOCKREADONLY
        RSDET.Open "SELECT * FROM CCOSTOS WHERE CODCCOSTO NOT IN (SELECT CODCCOSTO FROM TRABXCOSTO WHERE CODTRAB='" & OLDCOD & "')", DBSYSTEM, adOpenStatic, adLockReadOnly
        frmComun.CONECTAR RSDET
        frmComun.Show 1
        Dim BASICO As Single
        If VGUTIL(1) <> "" Then
            frValor.Show 1
            If VPTAREA = "0" Then
                If MsgBox("Desea agregar el registro con el Basico igual al Actual. Al cambiar el Basico en la Ficha del Trabajador, Automaticamente se Actualizara en este Centro de Costo", vbInformation) = vbYes Then
                    BASICO = Val(VPTAREA)
                Else
                    If RSTRABS.RecordCount > 0 Then RSTRABS.MOVE 0
                    Set RSDET = Nothing
                    Exit Sub
                End If
            Else
                BASICO = Val(VPTAREA)
            End If
            DBSYSTEM.Execute "INSERT INTO TRABXCOSTO VALUES ('" & OLDCOD & "','" & VGUTIL(1) & "'," & VPTAREA & ")"
        End If
        Set RSDET = Nothing
    End If
    If RSTRABS.RecordCount > 0 Then RSTRABS.MOVE 0
    REFRESCATRAB
End Sub

Private Sub CMAGREGAR_CLICK()
    If RSTRABS.EOF Then Exit Sub
    Set RSDET = Nothing
    RSDET.Open "SELECT * FROM CCOSTOS WHERE CODCCOSTO NOT IN (SELECT CODCCOSTO FROM TRABXCOSTO WHERE CODTRAB='" & RSTRABS!CODTRAB & "')", DBSYSTEM, adOpenStatic, adLockReadOnly
    frmComun.CONECTAR RSDET
    frmComun.Show 1
    Dim BASICO As Single
    If VGUTIL(1) <> "" Then
        frValor.Show 1
        If VPTAREA = "0" Then
            If MsgBox("Desea agregar el Registro con el Basico igual al actual. Al cambiar el Basico en la Ficha del Trabajador, Automaticamente se actualizara en este Centro de Costo", vbInformation) = vbYes Then
                BASICO = Val(VPTAREA)
            Else
                RSTRABS.MOVE 0
                Set RSDET = Nothing
                Exit Sub
            End If
        Else
            BASICO = Val(VPTAREA)
        End If
        DBSYSTEM.Execute "INSERT INTO TRABXCOSTO VALUES ('" & RSTRABS!CODTRAB & "','" & VGUTIL(1) & "'," & VPTAREA & ")"
    End If
    Set RSDET = Nothing
    RSTRABS.MOVE 0
End Sub
Private Sub cmCerrar_Click()
    Unload Me
End Sub

Private Sub CMIMPRIMIR_CLICK()
    'IMPRIMIR
    DBSTARPLAN.Execute "EXEC [TMP_TRBXCC] '" & REGSISTEMA.BASESQL & "'"
    With Reporte
        .Reset
        .Connect = "DSN=" & VGL_SERVERREP & ";User=sa;PWD=;DSQ=" & VGL_BASE & ""
        .ReportFileName = REGSISTEMA.REPORTES & "PLAN0057.RPT"
        .StoredProcParam(0) = REGSISTEMA.BASESQL
        .WindowTitle = "PLAN0057 - LISTADO DE TRABAJDADORES EN VARIOS CENTROS DE COSTOS"
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .WindowShowPrintBtn = True
        .WindowShowRefreshBtn = True
        .WindowShowSearchBtn = True
        .WindowShowPrintSetupBtn = True
        .Formulas(0) = "XEMPRESA='" & REGSISTEMA.EMPRESA & "'"
        If .Status <> 2 Then .Action = 1
        .Reset
    End With
End Sub

Private Sub CMQUITAR_CLICK()
    If MsgBox("Confirma que Desea quitar esta Asignación de Centro de Costo", vbYesNo + vbInformation) = vbNo Then Exit Sub
        If RSDET.RecordCount > 0 Then
            DBSYSTEM.Execute "DELETE FROM TRABXCOSTO WHERE CODTRAB='" & RSTRABS!CODTRAB & "' AND CODCCOSTO='" & RSDET!CODCCOSTO & "'"
        End If
    RSTRABS.MOVE 0
End Sub

Private Sub Form_Load()
    Set RSTRABS = New ADODB.Recordset
    RSTRABS.Open "SELECT DISTINCT A.CODTRAB,VWTRABAJ.NOMBRES AS NOMTRAB FROM TRABXCOSTO A, VWTRABAJ WHERE A.CODTRAB=VWTRABAJ.CODTRAB", DBSYSTEM, adOpenStatic, adLockReadOnly
    REFRESCATRAB
End Sub

Public Sub REFRESCATRAB()
    RSTRABS.Requery
    Set xTrabs.DataSource = RSTRABS
End Sub

Private Sub RSTRABS_MOVECOMPLETE(ByVal ADREASON As ADODB.EventReasonEnum, ByVal PERROR As ADODB.Error, ADSTATUS As ADODB.EventStatusEnum, ByVal PRECORDSET As ADODB.Recordset)
    If RSTRABS.EOF Or RSTRABS.RecordCount = 0 Then Exit Sub
    Set RSDET = Nothing
    RSDET.Open "SELECT TRABXCOSTO.CODCCOSTO, CCOSTOS.NOMBRE, TRABXCOSTO.BASICO FROM TRABXCOSTO, CCOSTOS WHERE TRABXCOSTO.CODCCOSTO=CCOSTOS.CODCCOSTO AND TRABXCOSTO.CODTRAB='" & RSTRABS!CODTRAB & "'", DBSYSTEM, adOpenStatic, adLockReadOnly
    Set xDetalles.DataSource = RSDET
End Sub

Private Sub XTRABS_HEADCLICK(ByVal COLINDEX As Integer)
    RSTRABS.Sort = xTrabs.Columns(COLINDEX).DataField
End Sub

