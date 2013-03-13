VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form FrmDeb 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Imprimir Debitos de Cuenta Corriente"
   ClientHeight    =   2580
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4875
   Icon            =   "FrmDeb.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   4875
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Detalle"
      Height          =   1275
      Left            =   60
      TabIndex        =   6
      Top             =   0
      Width           =   4740
      Begin VB.Label Label1 
         Caption         =   $"FrmDeb.frx":0442
         Height          =   780
         Left            =   735
         TabIndex        =   7
         Top             =   360
         Width           =   3945
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   120
         Picture         =   "FrmDeb.frx":04E7
         Top             =   390
         Width           =   240
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Opciones de Impresión"
      Height          =   1125
      Left            =   60
      TabIndex        =   2
      Top             =   1350
      Width           =   3555
      Begin VB.TextBox SqlCad 
         Height          =   300
         Left            =   2940
         TabIndex        =   5
         Top             =   795
         Visible         =   0   'False
         Width           =   645
      End
      Begin Crystal.CrystalReport Reporte 
         Left            =   3015
         Top             =   630
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.ComboBox cmbOp 
         Height          =   315
         ItemData        =   "FrmDeb.frx":0929
         Left            =   90
         List            =   "FrmDeb.frx":0933
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   255
         Width           =   3300
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Incluir Adelantos de Remuneraciones"
         Height          =   315
         Left            =   105
         TabIndex        =   3
         Top             =   720
         Width           =   3060
      End
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   390
      Left            =   3705
      TabIndex        =   1
      Top             =   2025
      Width           =   1065
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   390
      Left            =   3705
      TabIndex        =   0
      Top             =   1515
      Width           =   1065
   End
End
Attribute VB_Name = "FrmDeb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CMDACEPTAR_CLICK()
    Dim RUTPAGOSCTA As String, RUTMOVICTA As String
    Dim RUTCTAGRUPO As String, RUTNOMBOL As String
    Dim ANNOADEL As String, RUTADEL As String
    Dim INTO As String
    Screen.MousePointer = 11
    DBSTARPLAN.Execute "EXECUTE TMP_CREA_BOLETA '2000', '" & REGSISTEMA.BASESQL & "', " & Check1.Value & ",'" & VGL_COMPUTER & "'"
    With Reporte
        .Reset
        If cmbOp.ListIndex = 0 Then
            .ReportFileName = REGSISTEMA.REPORTES & "PLAN0020.RPT"
            .WindowTitle = "PLAN0020 - CONSOLIDADO DE SALDOS PENDIENTES POR PERIODO Y CONCEPTO DE CTA. CTE."
          Else:
            .ReportFileName = REGSISTEMA.REPORTES & "PLAN0021.RPT"
            .WindowTitle = "PLAN0021 - CONSOLIDADO DE SALDOS PENDIENTES POR CONCEPTO DE CTA. CTE. Y PERIODO "
        End If
        .Connect = "DSN=" & VGL_SERVERREP & ";User=sa;PWD=;DSQ=" & VGL_BASE & ""
        .StoredProcParam(0) = "2000"
        .StoredProcParam(1) = REGSISTEMA.BASESQL
        .StoredProcParam(2) = Check1.Value
        .StoredProcParam(3) = VGL_COMPUTER
        .Destination = crptToWindow
        .WindowShowPrintBtn = True
        .WindowShowRefreshBtn = True
        .WindowShowSearchBtn = True
        .WindowShowPrintSetupBtn = True
        .Formulas(0) = "XEMPRESA='" & REGSISTEMA.EMPRESA & "'"
        .Formulas(1) = "XRUC='" & REGSISTEMA.RUC & "'"
        If .Status <> 2 Then .Action = 1
    End With
    Screen.MousePointer = 1
End Sub

Private Sub CMDCANCELAR_CLICK()
    Unload Me
End Sub

Private Sub Form_Activate()
    cmbOp.ListIndex = 0
End Sub


