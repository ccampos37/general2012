VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frCTSCustodia 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Custodia de C.T.S."
   ClientHeight    =   5490
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8385
   Icon            =   "frCTSCustodia.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   8385
   StartUpPosition =   3  'Windows Default
   Begin Crystal.CrystalReport Reporte 
      Left            =   3435
      Top             =   2370
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Listado"
      Height          =   480
      Left            =   4020
      TabIndex        =   5
      Top             =   4785
      Width           =   1335
   End
   Begin VB.CommandButton cmAnular 
      Caption         =   "&Anular Custodia de C.T.S."
      Height          =   480
      Left            =   5460
      TabIndex        =   4
      Top             =   4785
      Width           =   1335
   End
   Begin VB.CommandButton cmCancelar 
      Caption         =   "&Efectuar el Pago de C.T.S."
      Height          =   480
      Left            =   1665
      TabIndex        =   3
      Top             =   4785
      Width           =   1335
   End
   Begin VB.CommandButton cmCustodia 
      Caption         =   "&Poner en Custodia"
      Height          =   480
      Left            =   225
      TabIndex        =   2
      Top             =   4785
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&Cerrar"
      Height          =   480
      Left            =   6900
      TabIndex        =   1
      Top             =   4785
      Width           =   1335
   End
   Begin MSDataGridLib.DataGrid xData 
      Height          =   4530
      Left            =   195
      TabIndex        =   0
      Top             =   165
      Width           =   8040
      _ExtentX        =   14182
      _ExtentY        =   7990
      _Version        =   393216
      AllowUpdate     =   -1  'True
      Appearance      =   0
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
      Caption         =   "Depósitos C.T.S. en Custodia por la Empresa"
      ColumnCount     =   4
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
         DataField       =   "Nombres"
         Caption         =   "Nombres"
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
         DataField       =   "ImporteCTS"
         Caption         =   "Importe Custodia C.T.S."
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
      BeginProperty Column03 
         DataField       =   "Nombre"
         Caption         =   "Periodo de Pago de C.T.S."
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
         BeginProperty Column00 
            ColumnWidth     =   794.835
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   3000.189
         EndProperty
         BeginProperty Column02 
            Alignment       =   1
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   3000.189
         EndProperty
      EndProperty
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   5295
      Left            =   90
      Top             =   75
      Width           =   8235
   End
End
Attribute VB_Name = "frCTSCustodia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RSCUSTO As New ADODB.Recordset

Private Sub CMANULAR_Click()
    If RSCUSTO.EOF Or RSCUSTO.RecordCount = 0 Then Exit Sub
    If MsgBox("SEGURO QUE DESEA CANCELAR EL INGRESO DE LA CUSTODIA DEL TRABAJADOR " & RSCUSTO!NOMBRES, vbQuestion + vbYesNo) = vbNo Then Exit Sub
    DBSYSTEM.Execute "UPDATE PLANCTS SET CUSTODIA=0 WHERE CODIGO=" & RSCUSTO!Codigo & " AND CODTRAB='" & RSCUSTO!CODTRAB & "'"
    REFRESCAR
End Sub

Private Sub CMCANCELAR_CLICK()
    If RSCUSTO.EOF Or RSCUSTO.RecordCount = 0 Then Exit Sub
    If MsgBox("DESEA EFECTUAR EL PAGO DE CTS EN CUSTODIA DEL TRABAJADOR: " & RSCUSTO!NOMBRES & " DEL PERIODO " & RSCUSTO!NOMBRE, vbYesNo + vbQuestion) = vbNo Then Exit Sub
    DBSYSTEM.Execute "UPDATE PLANCTS SET CUSTODIA=0, PAGOCUSTODIO=" & DateSQL(Date) & " WHERE CODIGO=" & RSCUSTO!Codigo & " AND CODTRAB='" & RSCUSTO!CODTRAB & "'"
    REFRESCAR
End Sub

Private Sub CMCUSTODIA_CLICK()
    Dim XCODE As Long
    Dim RSAUX As New ADODB.Recordset
    RSAUX.Open "SELECT CODIGO, NOMBRE FROM CTS", DBSYSTEM, adOpenKeyset, adLockReadOnly
    If RSAUX.EOF Then
        MsgBox "NO EXISTEN PERIODOS DE CTS REGISTRADOS EN EL SISTEMA", vbInformation
        Set RSAUX = Nothing
        Exit Sub
    End If
    frmComun.CONECTAR RSAUX
    frmComun.Show 1
    If VGUTIL(1) <> "" Then
        XCODE = VGUTIL(1)
        RSAUX.Close
        RSAUX.Open "SELECT CODTRAB, NOMBRES, IMPORTECTS FROM PLANCTS WHERE CUSTODIA<>1 AND CODIGO=" & VGUTIL(1), DBSYSTEM, adOpenKeyset, adLockOptimistic
        If RSAUX.EOF Then
            MsgBox "NO EXISTEN TRABAJADORES CON PAGOS DE CTS REGISTRADOS EN EL SISTEMA  PARA EL PERIODO SELECCIONADO", vbInformation
            Set RSAUX = Nothing
            Exit Sub
        End If
        frmComun.CONECTAR RSAUX
        frmComun.Show 1
        If VGUTIL(1) <> "" Then
            DBSYSTEM.Execute "UPDATE PLANCTS SET CUSTODIA=1 WHERE CODIGO=" & XCODE & " AND CODTRAB='" & VGUTIL(1) & "'"
            REFRESCAR
        End If
    End If
    Set RSAUX = Nothing
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
    If ExisteTablaAux(" [##TMPCTSCUSTODIA" & VGL_COMPUTER & "] ") Then DBSYSTEM.Execute "DROP TABLE  [##TMPCTSCUSTODIA" & VGL_COMPUTER & "] "
    DBSYSTEM.Execute "SELECT CTS.CODIGO, CODTRAB, NOMBRES, IMPORTECTS, NOMBRE INTO  [##TMPCTSCUSTODIA" & VGL_COMPUTER & "]  FROM CTS, PLANCTS WHERE CTS.CODIGO=PLANCTS.CODIGO AND CUSTODIA=1"
    With Reporte
        .WindowTitle = "PLAN0053 - LISTADO DE CTS EN CUSTODIA POR LA EMPRESA"
        .ReportFileName = REGSISTEMA.REPORTES & "PLAN0053.RPT"
        .Connect = "DSN=" & VGL_SERVER & ";User=sa;PWD=;DSQ=" & REGSISTEMA.BASESQL & ""
        .StoredProcParam(0) = " [##TMPCTSCUSTODIA" & VGL_COMPUTER & "] "
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .WindowShowPrintBtn = True
        .WindowShowRefreshBtn = True
        .WindowShowSearchBtn = True
        .WindowShowPrintSetupBtn = True
        .Formulas(0) = "XEMPRESA='" & REGSISTEMA.EMPRESA & "'"
        .Formulas(1) = "XRUC='RUC N° " & REGSISTEMA.RUC & "'"
        If .Status <> 2 Then .Action = 1
    End With
End Sub

Private Sub Form_Load()
    RSCUSTO.Open "SELECT CTS.CODIGO, CODTRAB, NOMBRES, IMPORTECTS, NOMBRE FROM CTS, PLANCTS WHERE CTS.CODIGO=PLANCTS.CODIGO AND CUSTODIA=1", DBSYSTEM, adOpenKeyset, adLockOptimistic
    REFRESCAR
End Sub

Private Sub FORM_UNLOAD(CANCEL As Integer)
    Set RSCUSTO = Nothing
End Sub

Public Sub REFRESCAR()
    RSCUSTO.Requery
    Set xData.DataSource = RSCUSTO
End Sub

