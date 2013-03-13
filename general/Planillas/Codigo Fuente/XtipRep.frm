VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form XtipRep 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tipo de Reporte"
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4935
   Icon            =   "XtipRep.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   4935
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   2715
      Left            =   75
      TabIndex        =   0
      Top             =   30
      Width           =   4740
      Begin Crystal.CrystalReport CtrReport 
         Left            =   300
         Top             =   1980
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Cancelar"
         Height          =   330
         Left            =   2970
         TabIndex        =   5
         Top             =   2280
         Width           =   1650
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Imprimir"
         Height          =   345
         Left            =   2970
         TabIndex        =   4
         Top             =   1860
         Width           =   1650
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Resumen por Centro de Costo"
         Height          =   345
         Index           =   2
         Left            =   120
         TabIndex        =   3
         Top             =   1230
         Width           =   2730
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Detallado por Centro de Costo"
         Height          =   345
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   765
         Width           =   2880
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Detallado por Trabajador"
         Height          =   345
         Index           =   0
         Left            =   150
         TabIndex        =   1
         Top             =   270
         Width           =   2625
      End
      Begin VB.Image Image1 
         Height          =   525
         Left            =   3555
         Picture         =   "XtipRep.frx":08CA
         Stretch         =   -1  'True
         Top             =   330
         Width           =   645
      End
   End
End
Attribute VB_Name = "XtipRep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Reporte As String
Public TITLE As String
Public MES As String
Dim OP As Integer
Private Sub Command1_Click()
Dim xTitulo As String
Dim Temporal As String
Screen.MousePointer = 11
With CtrReport
        .Reset
        Select Case OP
        Case 0
            .ReportFileName = REGSISTEMA.REPORTES & "PLANPR01.RPT"
            .WindowTitle = "PLANPR01.RPT" & TITLE
            xTitulo = "PROVISIONES MENSUALES DE " & UCase(TITLE) & " DEL MES DE " & UCase(MES)
            Temporal = " [##_TMPPROVISIONES" & VGL_COMPUTER & "] "
        Case 1
            .ReportFileName = REGSISTEMA.REPORTES & "PLANPR02.RPT"
            .WindowTitle = "PLANPR02.RPT" & TITLE
            xTitulo = "DETALLE POR CENTRO DE COSTO PROVISIONES DE " & UCase(TITLE) & " DEL MES DE " & UCase(MES)
            Temporal = " [##_TMPPROVCC" & VGL_COMPUTER & "] "
        Case 2
            .ReportFileName = REGSISTEMA.REPORTES & "PLANPR03.RPT"
            .WindowTitle = "PLANPR03.RPT" & TITLE
            xTitulo = "RESUMEN POR CENTRO DE COSTO PROVISIONES DE " & UCase(TITLE) & " DEL MES DE " & UCase(MES)
            Temporal = " [##_TMPPROVCC" & VGL_COMPUTER & "] "
        End Select
        .Connect = "DSN=" & VGL_SERVERREP & ";User=sa;PWD=;DSQ=planilla_pp"
        .StoredProcParam(0) = Temporal
        .Destination = crptToWindow
        .WindowShowPrintBtn = True
        .WindowShowRefreshBtn = True
        .WindowShowSearchBtn = True
        .WindowShowPrintSetupBtn = True
        .WindowTitle = xTitulo
        .Formulas(0) = "XEMPRESA='" & REGSISTEMA.EMPRESA & "'"
        .Formulas(1) = "XRUC='" & REGSISTEMA.RUC & "'"
        .Formulas(2) = "XTITULO='" & xTitulo & "'"
        If .Status <> 2 Then .Action = 1
End With
Screen.MousePointer = 1
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub OPTION1_CLICK(INDEX As Integer)
    OP = INDEX
End Sub

