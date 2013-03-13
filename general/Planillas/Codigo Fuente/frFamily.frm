VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C5FF36B7-A67E-11D3-9D9D-E6F193F7F854}#9.0#0"; "ApliCTxt.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frFamily 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Derechohabientes"
   ClientHeight    =   3645
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6210
   Icon            =   "frFamily.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   6210
   Tag             =   "Panel de Derechohabientes del Trabajador"
   Begin Crystal.CrystalReport Reporte 
      Left            =   2355
      Top             =   1440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSDataGridLib.DataGrid dgFamiliar 
      Height          =   2820
      Left            =   150
      TabIndex        =   0
      Top             =   675
      Visible         =   0   'False
      Width           =   5955
      _ExtentX        =   10504
      _ExtentY        =   4974
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
      Caption         =   "Derechohabientes de"
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
         MarqueeStyle    =   3
         AllowRowSizing  =   0   'False
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin AplisetControlText.Aplitext xTrab 
      Height          =   285
      Left            =   1125
      TabIndex        =   2
      Top             =   240
      Width           =   3405
      _ExtentX        =   6006
      _ExtentY        =   503
      Locked          =   -1  'True
      Text            =   ""
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Seleccione Trabajador"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1515
      TabIndex        =   3
      Top             =   1530
      Width           =   3195
   End
   Begin VB.Image Image2 
      Height          =   735
      Left            =   585
      Picture         =   "frFamily.frx":030A
      Stretch         =   -1  'True
      Top             =   1245
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Trabajador"
      Height          =   195
      Left            =   210
      TabIndex        =   1
      Top             =   285
      Width           =   765
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   5655
      Picture         =   "frFamily.frx":0BD4
      Top             =   60
      Width           =   480
   End
End
Attribute VB_Name = "frFamily"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RSFAM As New ADODB.Recordset
Dim REGACT As REGWIN
Dim ITSOPEN As Boolean

Private Sub DGFAMILIAR_DBLCLICK()
    COMANDOTOOLBAR "EDITAR"
End Sub

Private Sub Form_Activate()
    ActivarTools REGACT
End Sub

Private Sub Form_Load()
    ITSOPEN = False
    With REGACT
        .BUSCAR = True
        .EDITAR = True
        .ELIMINAR = True
        .FILTRAR = False
        .IMPRIMIR = True
        .NUEVO = True
        .PRELIMINAR = True
    End With
End Sub

Private Sub FORM_UNLOAD(CANCEL As Integer)
    Set RSFAM = Nothing
End Sub

Private Sub XTRAB_DBLCLICK()
    Dim RSTRAB As New ADODB.Recordset
    RSTRAB.Open "SELECT CODTRAB, NOMBRES FROM VWTRABAJ", DBSYSTEM, adOpenKeyset, adLockReadOnly
    If RSTRAB.EOF Or RSTRAB.RecordCount = 0 Then
        MsgBox "No se han encontrado registro de trabajadores", vbCritical
        Set RSTRAB = Nothing
        Exit Sub
    End If
    frmComun.CONECTAR RSTRAB
    frmComun.Show 1
    If VGUTIL(1) <> "" Then
        xTrab.Tag = RSTRAB!CODTRAB
        xTrab.Text = RSTRAB!CODTRAB & " : " & RSTRAB!NOMBRES
        MOSTRARFAMILY xTrab.Tag
    End If
    Set RSTRAB = Nothing
End Sub

Public Sub MOSTRARFAMILY(Codigo As String)
    If ITSOPEN Then RSFAM.Close
    RSFAM.Open "SELECT CODDER, APEPAT + ' ' + APEMAT + ' ' + NOMBRE AS NOMBRES, VINCULO, FECHANAC FROM FAMILIAR WHERE CODTRAB='" & Codigo & "' ORDER BY VINCULO", DBSYSTEM, adOpenKeyset
    Set dgFamiliar.DataSource = RSFAM
    If RSFAM.RecordCount = 0 Then
        dgFamiliar.Visible = False
    Else
        dgFamiliar.Visible = True
        FORMATEARDG
    End If
    ITSOPEN = True
End Sub

Public Sub COMANDOTOOLBAR(COMANDO As String)
    Select Case UCase(COMANDO)
        Case "NUEVO"
            If xTrab.Tag = "" Then
                MsgBox "No es posible agregar un derechohabiente sin un trabajador seleccionado", vbCritical
                Exit Sub
            End If
            VPTAREA = "NUEVO"
            frEdFam.Show 1
            RSFAM.Requery
        Case "EDITAR"
            If RSFAM.State = 0 Then Exit Sub
            If RSFAM.RecordCount = 0 Then Exit Sub
            If RSFAM.EOF Then Exit Sub
            VPTAREA = "" & RSFAM!CODDER
            frEdFam.Show 1
            RSFAM.Requery
        Case "ELIMINAR"
            If RSFAM.State = 0 Then Exit Sub
            If RSFAM.RecordCount = 0 Then Exit Sub
            If MsgBox("Desea eliminar el registro del derechohabiente seleccionado", vbYesNo + vbQuestion) = vbNo Then Exit Sub
            DBSYSTEM.Execute "DELETE FROM FAMILIAR WHERE CODDER=" & RSFAM!CODDER
            RSFAM.Requery
        Case "IMPRIMIR"
            If RSFAM.State = 0 Then Exit Sub
            If RSFAM.RecordCount = 0 Then Exit Sub
            With Reporte
                .WindowTitle = "LISTADO DE DERECHOHABIENTOS POR TRABAJADOR"
                .ReportFileName = REGSISTEMA.REPORTES & "PLAN0022.RPT"
                .Connect = "DSN=" & VGL_SERVER & ";User=sa;PWD=;DSQ=" & REGSISTEMA.BASESQL & ""
                .StoredProcParam(0) = xTrab.Tag
                .StoredProcParam(1) = REGSISTEMA.BASESQL
                .Destination = crptToWindow
                .WindowState = crptMaximized
                .WindowShowPrintBtn = True
                .WindowShowRefreshBtn = True
                .WindowShowSearchBtn = True
                .WindowShowPrintSetupBtn = True
                .Formulas(0) = "XEMPRESA='" & REGSISTEMA.EMPRESA & "'"
                .Formulas(1) = "XRUC='RUC N° " & REGSISTEMA.RUC & "'"
                .Formulas(2) = "''"
                If .Status <> 2 Then .Action = 1
            End With
    End Select
End Sub

Public Sub FORMATEARDG()
    With dgFamiliar
        .Columns("CODDER").Visible = False
    End With
End Sub

