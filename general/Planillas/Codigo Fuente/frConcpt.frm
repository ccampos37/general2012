VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frConcpt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Conceptos de Remuneraciones"
   ClientHeight    =   5550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8490
   Icon            =   "frConcpt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   8490
   Tag             =   "Panel de Conceptos Remuneraciones"
   Begin Crystal.CrystalReport RptConceptos 
      Left            =   375
      Top             =   5085
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Comprobar integridad"
      Height          =   375
      Left            =   6465
      TabIndex        =   2
      Top             =   5070
      Width           =   1905
   End
   Begin MSDataGridLib.DataGrid dgConceptos 
      Height          =   4035
      Left            =   135
      TabIndex        =   0
      Top             =   930
      Width           =   8235
      _ExtentX        =   14526
      _ExtentY        =   7117
      _Version        =   393216
      AllowUpdate     =   0   'False
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
      Caption         =   "Conceptos de Remuneraciones"
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
         AllowRowSizing  =   0   'False
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frConcpt.frx":08CA
      ForeColor       =   &H8000000E&
      Height          =   630
      Left            =   780
      TabIndex        =   1
      Top             =   150
      Width           =   7515
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   135
      Picture         =   "frConcpt.frx":0987
      Top             =   195
      Width           =   240
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   225
      Picture         =   "frConcpt.frx":0CC9
      Top             =   210
      Width           =   480
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000D&
      BackStyle       =   1  'Opaque
      Height          =   840
      Left            =   -15
      Top             =   0
      Width           =   8520
   End
End
Attribute VB_Name = "frConcpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RSRUBROS As New ADODB.Recordset
Dim REGACT As REGWIN
'TENEMOS QUE HABRIRLO CON EVENTOS USAREMOS WITHEVENTS

Private Sub Command1_Click()
    If RSRUBROS.EOF Then Exit Sub
    RSRUBROS.MoveFirst
    Dim X As Long, Z As Byte
    X = 0
    Do While Not RSRUBROS.EOF
        If Trim(RSRUBROS!COLPLANILLA) <> "" Then
            DBSYSTEM.Execute "UPDATE COLUMPL SET TIPO=TIPO WHERE CODIGO='" & Trim(RSRUBROS!COLPLANILLA) & "'", X
            If X = 0 Then
                Z = MsgBox("El concepto de remuneracion " & RSRUBROS!NOMBRE & " PRESENTA COMO COLUMNA DE PLANILLA EL CÓDIGO " & RSRUBROS!COLPLANILLA & " EL CUAL NO EXISTE DENTRO DE LA BASE DE DATOS. DESEA DEPURAR EL CONCEPTO DE REMUNERACIÓN", vbQuestion + vbYesNoCancel)
                If Z = vbCancel Then Exit Sub
                If Z = vbYes Then
                    VPTAREA = "EDITAR"
                    VPCODTMP = RSRUBROS!Codigo
                    frECnpt.Show 1
                End If
            End If
        End If
        RSRUBROS.MoveNext
    Loop
    RSRUBROS.Requery
    FORMATEARDG
End Sub

Private Sub DGCONCEPTOS_DBLCLICK()
    Dim BOOK As Variant
    If RSRUBROS.RecordCount = 0 Then Exit Sub
    BOOK = RSRUBROS.Bookmark
    VPTAREA = "EDITAR"
    VPCODTMP = RSRUBROS!Codigo
    frECnpt.Show 1
    RSRUBROS.Requery
    FORMATEARDG
    RSRUBROS.Bookmark = BOOK
End Sub

Private Sub DGCONCEPTOS_HEADCLICK(ByVal COLINDEX As Integer)
    RSRUBROS.Sort = dgConceptos.Columns(COLINDEX).Caption
End Sub

Private Sub Form_Activate()
    ActivarTools REGACT
End Sub

Private Sub Form_Load()
   RSRUBROS.Open "CONCEPTOS", DBSYSTEM, adOpenKeyset, adLockOptimistic
    Set dgConceptos.DataSource = RSRUBROS
    FORMATEARDG
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

Public Sub COMANDOTOOLBAR(COMANDO As String)
Dim X As Integer
    Select Case UCase(COMANDO)
        Case "NUEVO"
            VPTAREA = "NUEVO"
            frECnpt.Show 1
            RSRUBROS.Requery
            FORMATEARDG
        Case "EDITAR"
            If RSRUBROS.EOF Then Exit Sub
            VPTAREA = "EDITAR"
            VPCODTMP = RSRUBROS!Codigo
            frECnpt.Show 1
            RSRUBROS.Requery
            FORMATEARDG
        Case "ELIMINAR"
            If RSRUBROS.EOF Then Exit Sub
            If RSRUBROS!FLAG = 1 Then
                MsgBox "El rubro no se puede eliminar pues es considerado como rubro de sistema. El usuario no puede eliminar los rubros para vacaciones, gratificaciones o quinta categoria y los conceptos de cuentas", vbCritical
                MsgBox "Se abortara del proceso de eliminacion de registros", vbInformation
                Exit Sub
            End If
            If MsgBox("Realmente desea eliminar el registro seleccionado", vbYesNo + vbQuestion) = vbNo Then Exit Sub
            Dim RSAX As New ADODB.Recordset
            RSAX.Open "SELECT CONCEPTO FROM FORMARUBS WHERE CONCEPTO='" & RSRUBROS!Codigo & "'", DBSYSTEM, adOpenStatic
            If RSAX.EOF Then
                RSRUBROS.Delete
            Else
                MsgBox "No se puede eliminar el concepto de remuneraciones, pues se encuentra registrado dentro de uno o mas formatos de boleta", vbCritical
            End If
            Set RSAX = Nothing
        Case "IMPRIMIR"
            Screen.MousePointer = 11
            If ExisteTablaAux(" [##TMPCONCEPTOS" & VGL_COMPUTER & "] ") Then DBSYSTEM.Execute "DROP TABLE  [##TMPCONCEPTOS" & VGL_COMPUTER & "] "
                DBSYSTEM.Execute "SELECT * INTO  [##TMPCONCEPTOS" & VGL_COMPUTER & "]  FROM CONCEPTOS", X
                With RptConceptos
                    .Reset
                    .WindowTitle = "PLAN0088 - REPORTE DE TRABAJADORES POR TIPO"
                    .ReportFileName = REGSISTEMA.REPORTES & "PLAN0088.RPT"
                    .Connect = "DSN=" & VGL_SERVERREP & ";User=sa;PWD=;DSQ=" & VGL_BASE & ""
                    .StoredProcParam(0) = "[##TMPCONCEPTOS" & VGL_COMPUTER & "]"
                    .Destination = crptToWindow
                    .WindowState = crptMaximized
                    .WindowShowPrintBtn = True
                    .WindowShowRefreshBtn = True
                    .WindowShowSearchBtn = True
                    .WindowShowPrintSetupBtn = True
                    .Formulas(0) = "EMPRESA='" & REGSISTEMA.EMPRESA & "'"
                    .Formulas(1) = "RUC='" & REGSISTEMA.RUC & "'"
                    If .Status <> 2 Then .Action = 1
                End With
            Screen.MousePointer = 1
        Case "PRELIMINAR"
            Screen.MousePointer = 11
            If ExisteTablaAux(" [##TMPCONCEPTOS" & VGL_COMPUTER & "] ") Then DBSYSTEM.Execute "DROP TABLE  [##TMPCONCEPTOS" & VGL_COMPUTER & "] "
                DBSYSTEM.Execute "SELECT * INTO  [##TMPCONCEPTOS" & VGL_COMPUTER & "]  FROM CONCEPTOS", X
                With RptConceptos
                    .Reset
                    .WindowTitle = "PLAN0088 - REPORTE DE TRABAJADORES POR TIPO"
                    .ReportFileName = REGSISTEMA.REPORTES & "PLAN0088.RPT"
                    .Connect = "DSN=" & VGL_SERVER & ";User=sa;PWD=;DSQ=" & REGSISTEMA.BASESQL & ""
                    .StoredProcParam(0) = " [##TMPCONCEPTOS" & VGL_COMPUTER & "] "
                    .Destination = crptToWindow
                    .WindowState = crptMaximized
                    .WindowShowPrintBtn = True
                    .WindowShowRefreshBtn = True
                    .WindowShowSearchBtn = True
                    .WindowShowPrintSetupBtn = True
                    .Formulas(0) = "EMPRESA='" & REGSISTEMA.EMPRESA & "'"
                    .Formulas(1) = "RUC='" & REGSISTEMA.RUC & "'"
                    If .Status <> 2 Then .Action = 1
                End With
            Screen.MousePointer = 1
    End Select
End Sub

Private Sub FORM_UNLOAD(CANCEL As Integer)
    Set RSRUBROS = Nothing
End Sub

Public Sub FORMATEARDG()
    With dgConceptos
        .Columns("FORMULA").Width = 6300
        .Columns("NOMBRE").Width = 3000
        .Columns("TIPO").Visible = True
        .Columns("ESESCRITO").Visible = False
        .Columns("MONEDA").Visible = False
        .Columns("TIPOQUINTA").Visible = False
        .Columns("TIPOREMU").Visible = False
        .Columns("SUMAAFP").Visible = False
        .Columns("SUMASALUD").Visible = False
        .Columns("SUMAIES").Visible = False
        .Columns("SUMARENTA").Visible = False
        .Columns("SUMASCTR").Visible = False
        .Columns("SUMACTS").Visible = False
        .Columns("SUMAVAC").Visible = False
        .Columns("SUMAGRAT").Visible = False
        .Columns("SUMAT1").Visible = False
        .Columns("SUMAT2").Visible = False
        .Columns("SUMAT3").Visible = False
        .Columns("SUMAT4").Visible = False
        .Columns("SUMAT5").Visible = False
    End With
End Sub

