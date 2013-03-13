VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmBilletes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Billetaje de Remuneraciones"
   ClientHeight    =   5220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frBilletes.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin Crystal.CrystalReport Reporte 
      Left            =   1590
      Top             =   2235
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton Command3 
      Cancel          =   -1  'True
      Caption         =   "&Cerrar"
      Height          =   375
      Left            =   3180
      TabIndex        =   4
      Top             =   4530
      Width           =   1335
   End
   Begin VB.CommandButton cmImprimir 
      Caption         =   "&Imprimir"
      Height          =   375
      Left            =   3180
      TabIndex        =   3
      Top             =   4005
      Width           =   1335
   End
   Begin VB.CommandButton cmRecalcular 
      Caption         =   "&Recalcular"
      Height          =   375
      Left            =   3180
      TabIndex        =   2
      Top             =   3480
      Width           =   1335
   End
   Begin MSDataGridLib.DataGrid DGResult 
      Height          =   3015
      Left            =   210
      TabIndex        =   0
      Top             =   150
      Width           =   4305
      _ExtentX        =   7594
      _ExtentY        =   5318
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
      Caption         =   "Billetaje de Remuneraciones"
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
         MarqueeStyle    =   4
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DGBilletes 
      Height          =   1785
      Left            =   210
      TabIndex        =   1
      Top             =   3285
      Width           =   2280
      _ExtentX        =   4022
      _ExtentY        =   3149
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
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
      Caption         =   "Monedas"
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
End
Attribute VB_Name = "frmBilletes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RSBILLETES As ADODB.Recordset
Dim RSRESULT As ADODB.Recordset

Private Sub CMIMPRIMIR_CLICK()
    DBSTARPLAN.Execute "EXECUTE [ASISTMP] ' [##BILLETAJE" & VGL_COMPUTER & "] '"
    With Reporte
        .WindowTitle = "REPORTE DE DISTRIBUCIÓN MONETARIA"
        .ReportFileName = REGSISTEMA.REPORTES & "PLAN0007.RPT"
        .Connect = "DSN=" & VGL_SERVERREP & ";User=sa;PWD=;DSQ=" & VGL_BASE & ""
        .StoredProcParam(0) = " [##BILLETAJE" & VGL_COMPUTER & "] "
        .Destination = crptToWindow
        .WindowState = crptNormal
        .WindowShowPrintBtn = True
        .WindowShowRefreshBtn = True
        .WindowShowSearchBtn = True
        .WindowShowPrintSetupBtn = True
        .Formulas(0) = "XEMPRESA='" & REGSISTEMA.EMPRESA & "'"
        '.Formulas(1) = "XMES='CORRESPONDIENTE A: " & frBolEmit.Lista.SelectedItem.Text & "'"
        If .Status <> 2 Then .Action = 1
    End With
End Sub

Private Sub CMRECALCULAR_Click()
    CambiaPanelBD True
    On Error Resume Next
    RSBILLETES.Requery
    Set DGBilletes.DataSource = RSBILLETES
    If ExisteTablaAux(" [##BILLETAJE" & VGL_COMPUTER & "] ") Then
        DBSYSTEM.Execute "DROP TABLE  [##BILLETAJE" & VGL_COMPUTER & "] "
    End If
    DBSYSTEM.Execute "CREATE TABLE  [##BILLETAJE" & VGL_COMPUTER & "]  (BILLETE  Numeric(20,2) , CANTIDAD INT)"
    RSBILLETES.MoveFirst
    Do While Not RSBILLETES.EOF
        DBSYSTEM.Execute "INSERT INTO  [##BILLETAJE" & VGL_COMPUTER & "]  (BILLETE, CANTIDAD) VALUES (" & RSBILLETES!BILLETE & ",0)"
        RSBILLETES.MoveNext
    Loop
    Set RSRESULT = Nothing
    Set RSRESULT = New ADODB.Recordset
    If VPTAREA = "ADELANTO" Then
        'PROVENIENTE DESDE FRADELEMIT
        RSRESULT.Open "SELECT * FROM  [##_TMPLSTADEL" & VGL_COMPUTER & "] ", DBSYSTEM, adOpenStatic
    Else
        'VPTAREA="BOLETAS". PROVENIENTE DESDE FRBOLEMIT
        RSRESULT.Open "SELECT * FROM  [##_TMPLSTBOL" & VGL_COMPUTER & "] ", DBSYSTEM, adOpenStatic
    End If
    Dim xValor As Single, XRESTO As Single, XDIV As Integer
    Do While Not RSRESULT.EOF
        xValor = Round(RSRESULT!Neto, 2)
        RSBILLETES.MoveFirst
        Do While Not RSBILLETES.EOF
            If xValor < 1 Then
                XDIV = (xValor * 10) \ (RSBILLETES!BILLETE * 10)
            Else
                XDIV = Int(xValor) \ RSBILLETES!BILLETE
            End If
            xValor = xValor - (XDIV * RSBILLETES!BILLETE)
            If XDIV > 0 Then
                DBSYSTEM.Execute "UPDATE  [##BILLETAJE" & VGL_COMPUTER & "]  SET CANTIDAD=CANTIDAD+" & XDIV & " WHERE BILLETE= " & RSBILLETES!BILLETE
            End If
            If xValor = 0 Then Exit Do
            RSBILLETES.MoveNext
        Loop
        RSRESULT.MoveNext
    Loop
    Set RSRESULT = Nothing
    Set RSRESULT = New ADODB.Recordset
    RSRESULT.Open "SELECT BILLETE, CANTIDAD, (BILLETE*CANTIDAD) AS TOTAL FROM  [##BILLETAJE" & VGL_COMPUTER & "] ", DBSYSTEM, adOpenStatic
    Set DGResult.DataSource = RSRESULT
    DGResult.Columns("BILLETE").Alignment = dbgCenter
    DGResult.Columns("CANTIDAD").Alignment = dbgCenter
    DGResult.Columns("TOTAL").Alignment = dbgRight
    DGResult.Columns("TOTAL").NumberFormat = "##,##0.00 "
    CambiaPanelBD False
End Sub

Private Sub Command3_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Set RSBILLETES = New ADODB.Recordset
    Set RSRESULT = New ADODB.Recordset
    RSBILLETES.Open "SELECT * FROM BILLETES ORDER BY BILLETE DESC", DBSYSTEM, adOpenKeyset, adLockOptimistic
    Set DGBilletes.DataSource = RSBILLETES
    DGBilletes.Columns("BILLETE").Width = 1700
    CMRECALCULAR_Click
End Sub

Private Sub FORM_UNLOAD(CANCEL As Integer)
    Set RSBILLETES = Nothing
    Set RSRESULT = Nothing
End Sub

