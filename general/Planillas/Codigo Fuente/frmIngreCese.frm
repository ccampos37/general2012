VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmIngreCese 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reporte de Ingresos/Egresos"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5070
   Icon            =   "frmIngreCese.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   5070
   Begin Crystal.CrystalReport Reporte 
      Left            =   1785
      Top             =   2025
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmSalir 
      Caption         =   "&Salir"
      Height          =   360
      Left            =   2738
      TabIndex        =   11
      Top             =   4245
      Width           =   1440
   End
   Begin VB.CommandButton cmImprimir 
      Caption         =   "&Imprimir"
      Height          =   360
      Left            =   893
      TabIndex        =   10
      Top             =   4245
      Width           =   1440
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tipo de Reporte"
      Height          =   765
      Left            =   120
      TabIndex        =   7
      Top             =   75
      Width           =   2880
      Begin VB.OptionButton Option2 
         Caption         =   "Egresos"
         Height          =   195
         Left            =   1620
         TabIndex        =   9
         Top             =   375
         Width           =   1020
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Ingresos"
         Height          =   195
         Left            =   210
         TabIndex        =   8
         Top             =   375
         Value           =   -1  'True
         Width           =   1020
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   165
      Top             =   2640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIngreCese.frx":08CA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Selección de Registros"
      Height          =   855
      Left            =   86
      TabIndex        =   0
      Top             =   915
      Width           =   4860
      Begin MSComCtl2.DTPicker xFechaFin 
         Height          =   300
         Left            =   3255
         TabIndex        =   4
         ToolTipText     =   "Fecha de finalización del Reporte"
         Top             =   375
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   529
         _Version        =   393216
         Format          =   62193665
         CurrentDate     =   36930
      End
      Begin MSComCtl2.DTPicker xFechaIni 
         Height          =   300
         Left            =   1035
         TabIndex        =   3
         ToolTipText     =   "Fecha de Inicio del reporte"
         Top             =   360
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   529
         _Version        =   393216
         Format          =   62193665
         CurrentDate     =   36930
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         Height          =   195
         Left            =   2640
         TabIndex        =   2
         Top             =   435
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         Height          =   195
         Left            =   225
         TabIndex        =   1
         Top             =   413
         Width           =   465
      End
   End
   Begin MSComctlLib.ListView LEmpresas 
      Height          =   1965
      Left            =   120
      TabIndex        =   5
      Top             =   2100
      Width           =   4860
      _ExtentX        =   8573
      _ExtentY        =   3466
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID Empresa"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nombre"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Ruta"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   4050
      Picture         =   "frmIngreCese.frx":0C1E
      Top             =   345
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   4305
      Picture         =   "frmIngreCese.frx":14E8
      Stretch         =   -1  'True
      Top             =   195
      Width           =   465
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Selección de Empresas"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   1890
      Width           =   1665
   End
End
Attribute VB_Name = "frmIngreCese"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub CMIMPRIMIR_CLICK()
    CambiaPanelBD True
    If ExisteTablaAux("[##TMPINGEGR" & VGL_COMPUTER & "]") Then DBSYSTEM.Execute "DROP TABLE  [##TMPINGEGR" & VGL_COMPUTER & "]"
    DBSYSTEM.Execute "CREATE TABLE  [##TMPINGEGR" & VGL_COMPUTER & "] (RUC VARCHAR(11), EMPRESA VARCHAR(70), CODTRAB VARCHAR(8), TRABAJADOR VARCHAR(100), DOCUMENTO VARCHAR(15), FECHAING DATETIME, CARGO VARCHAR(40), BASICO  Numeric(20,2) , FECHACESE DATETIME, CODCOSTO VARCHAR(10), CENTROCOSTO VARCHAR(80))"
    Dim RsEmp As New ADODB.Recordset
    Dim XITEM As ListItem
    For Each XITEM In LEmpresas.ListItems
        If XITEM.Checked Then
                If Option1.Value Then
                    DBSYSTEM.Execute "INSERT INTO  [##TMPINGEGR" & VGL_COMPUTER & "] SELECT '" & XITEM.Text & "' AS RUC, '" & XITEM.SubItems(1) & "' AS EMPRESA, CODTRAB, NOMBRES AS TRABAJADOR, DOCIDEN AS DOCUMENTO, FECHAING, CARGO, BASICO, FECHACESE, CODCCOSTO, CENTRO AS CENTROCOSTO FROM " & XITEM.SubItems(2) & ".dbo.VWTRABAJ WHERE FECHAING BETWEEN " & DateSQL(xFechaIni.Value) & " AND " & DateSQL(xFechaFin.Value)
                Else
                    DBSYSTEM.Execute "INSERT INTO  [##TMPINGEGR" & VGL_COMPUTER & "] SELECT '" & XITEM.Text & "' AS RUC, '" & XITEM.SubItems(1) & "' AS EMPRESA, CODTRAB, NOMBRES AS TRABAJADOR, DOCIDEN AS DOCUMENTO, FECHAING, CARGO, BASICO, FECHACESE, CODCCOSTO, CENTRO AS CENTROCOSTO FROM " & XITEM.SubItems(2) & ".dbo.VWTRABAJ WHERE FECHACESE BETWEEN " & DateSQL(xFechaIni.Value) & " AND " & DateSQL(xFechaFin.Value)
                End If
        End If
    Next
    DBSYSTEM.Execute "UPDATE  [##TMPINGEGR" & VGL_COMPUTER & "] SET CODTRAB=CODTRAB"
    With Reporte
        .Reset
        If Option1.Value Then
            .WindowTitle = "PLAN0086 - Reporte de Ingresos"
            .ReportFileName = REGSISTEMA.REPORTES & "PLAN0086.RPT"
        Else
            .WindowTitle = "PLAN0087 - Reporte de Egresos"
            .ReportFileName = REGSISTEMA.REPORTES & "PLAN0087.RPT"
        End If
        .Connect = "DSN=" & VGL_SERVERREP & ";User=sa;PWD=;DSQ=planilla_pp"
        .StoredProcParam(0) = " [##TMPINGEGR" & VGL_COMPUTER & "]"
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .WindowShowPrintBtn = True
        .WindowShowRefreshBtn = True
        .WindowShowSearchBtn = True
        .WindowShowPrintSetupBtn = True
        .Formulas(0) = "xFechaIni='" & xFechaIni.Value & "'"
        .Formulas(1) = "xFechaFin='" & xFechaFin.Value & "'"
        If .Status <> 2 Then .Action = 1
    End With
    CambiaPanelBD False
End Sub

Private Sub CMSALIR_CLICK()
    Unload Me
End Sub

Private Sub Form_Load()
    CambiaPanelBD True
    CargaEmp
    xFechaIni.Value = Date
    xFechaFin.Value = Date
    CambiaPanelBD False
End Sub

Public Sub CargaEmp()
    Dim RsEmp As New ADODB.Recordset
    Dim XITEM As ListItem
    Set RsEmp = Nothing
    RsEmp.Open "SELECT * FROM EMPRESAS ORDER BY NOMBRE", DBSTARPLAN, adOpenStatic, adLockOptimistic
    RsEmp.Requery
    LEmpresas.ListItems.Clear
    Do While Not RsEmp.EOF
        Set XITEM = LEmpresas.ListItems.Add(, "R" & RsEmp!RUC, RsEmp!RUC, , 1)
        XITEM.SubItems(1) = RsEmp!NOMBRE
        XITEM.SubItems(2) = RsEmp!DIRALMACEN
        RsEmp.MoveNext
    Loop
    RsEmp.Close
    LEmpresas.ColumnHeaders(3).Width = 0
    LEmpresas.Refresh
End Sub
