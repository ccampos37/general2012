VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form RptNotaCreRe 
   Caption         =   "Reporte de Notas"
   ClientHeight    =   3300
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4995
   LinkTopic       =   "Form1"
   ScaleHeight     =   3300
   ScaleWidth      =   4995
   StartUpPosition =   3  'Windows Default
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   120
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame1 
      Height          =   1905
      Left            =   480
      TabIndex        =   2
      Top             =   480
      Width           =   3660
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   360
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker DTP_FechaInicio 
         Height          =   330
         Left            =   1665
         TabIndex        =   3
         Top             =   885
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         _Version        =   393216
         Format          =   62193665
         CurrentDate     =   37586
      End
      Begin MSComCtl2.DTPicker DTP_FechaFin 
         Height          =   330
         Left            =   1665
         TabIndex        =   4
         Top             =   1380
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         _Version        =   393216
         Format          =   62193665
         CurrentDate     =   37586
      End
      Begin VB.Label Label2 
         Caption         =   "Tipo :"
         Height          =   255
         Left            =   1200
         TabIndex        =   8
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Hasta la Fecha :"
         Height          =   255
         Index           =   2
         Left            =   390
         TabIndex        =   6
         Top             =   1440
         Width           =   1245
      End
      Begin VB.Label Label1 
         Caption         =   "Desde la Fecha :"
         Height          =   255
         Index           =   0
         Left            =   390
         TabIndex        =   5
         Top             =   945
         Width           =   1245
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Top             =   2640
      Width           =   1440
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   2580
      TabIndex        =   0
      Top             =   2640
      Width           =   1380
   End
End
Attribute VB_Name = "RptNotaCreRe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim adll As New dllgeneral.dll_general
Dim titu As String
Private Sub cmdAceptar_Click()
    'Call ImprimirReport
Dim busca As New dll_apis
Dim i As Integer

On Error GoTo X
    Screen.MousePointer = 11
    'With FrmMenu.oCrystalReport
    With CrystalReport1
        .Reset
        .ReportFileName = VGParamSistem.Rutareport & "vt_RepNotCre.rpt"
        .Connect = VGcadenareport2
        .DiscardSavedData = True
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .WindowShowPrintSetupBtn = True
        .WindowShowExportBtn = True
        .WindowShowZoomCtl = True
        .WindowShowNavigationCtls = True
        .WindowShowPrintBtn = True
        .WindowTitle = "Informe de Notas de Credito por Concepto"
        
        .formulas(0) = "emp ='" & VGParametros.nomempresa & "'"
        
        Select Case Left(Trim(Combo1.Text), 1)
            Case "A"
                .formulas(1) = "titulo ='" & "REPORTE DE NOTAS DE ABONO" & "'"
            Case "C"
                .formulas(1) = "titulo ='" & "REPORTE DE NOTAS DE CARGO" & "'"
            Case Else    'case  A"
                .formulas(1) = "titulo ='" & " REPORTE DE NOTAS DE CARGO/ABONO " & "'"
        End Select
        
        .StoredProcParam(0) = CStr(VGCNx.DefaultDatabase)
        .StoredProcParam(1) = VGParametros.empresacodigo
        .StoredProcParam(2) = CStr(DTP_FechaInicio.Value)
        .StoredProcParam(3) = CStr(DTP_FechaFin.Value)
        .StoredProcParam(4) = Trim(adll.ComboDato(Trim(Combo1.Text)))

        If .Status <> 2 Then .Action = 1
        
    End With
    Screen.MousePointer = 1
    Exit Sub
X:
  If Err.Number = 9 Then Resume Next
  Screen.MousePointer = 1
  MsgBox "Error inesperado: " & Err.Number & "  " & Err.Description, vbExclamation
End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub Form_Load()
   MostrarForm Me, "C2"
   DTP_FechaInicio.Value = "01/" & Format(Month(Now), "00") & "/" & Year(Date)
   DTP_FechaFin.Value = Format(Date, "dd/mm/yyyy")
   'Ctr_Cliente.conexion cn
   
  Combo1.Clear
  Combo1.AddItem "A - ABONO"
  Combo1.AddItem "C - CARGO"
  Combo1.AddItem "T - TODOS"
  
  Combo1.ListIndex = 0
End Sub


Sub ImprimirReport()
'Dim arrform(1) As Variant, arrparm(4) As Variant
'Dim NombreRep As String, CadOrden As String
'Dim NombrePC As String
'Dim mon As String
'    Randomize   ' Inicializa el generador de números aleatorios.
'    NombrePC = Trim(Str(CLng(Rnd * 10000000)))
'    arrparm(0) = VGcnx.DefaultDatabase
'    'arrparm(1) = Format(DTPicker1.Value, "dd/mm/yyyy")
'    arrparm(1) = Format(DTP_FechaInicio.Value, "dd/mm/yyyy")
'    arrparm(2) = Format(DTP_FechaFin.Value, "dd/mm/yyyy")
'    arrparm(3) = Left(Trim(Combo1.Text), 2)
'
'
'
'    'arrform(0) = "@Fecha='" & Format(DTP_Fecha.Value, "dd/mm/yyyy") & "'"
'    arrform(1) = "@emp ='" & VGNemp & "'"
'    NombreRep = "REPNOTCRE.rpt"
'    CadOrden = ""
'
'
'    Call ImpresionRptProc(NombreRep, arrform, arrparm, CadOrden, "Reporte de Notas ")
End Sub

Sub ImpresionRptProc(cNombreReporte As String, PFormulas(), Param(), Optional ORDEN As String, Optional Titulo As String)
Dim strBuscar As New dll_apis
Dim i As Integer

On Error GoTo X
    Screen.MousePointer = 11
    
    With CrystalReport1
        .Reset
        
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .WindowTitle = "Reporte de Notas"
        .ReportFileName = VGParamSistem.Rutareport & "RepNotCre.rpt"
        .LogOnServer "pdssql.dll", _
         strBuscar.LeerIni(App.Path & "\MARFICE.INI", "BDGENERAL", "SERVIDOR", ""), _
         strBuscar.LeerIni(App.Path & "\MARFICE.INI", "BDGENERAL", "BDDATOS", ""), _
         strBuscar.LeerIni(App.Path & "\MARFICE.INI", "BDGENERAL", "USUARIO", ""), _
         strBuscar.LeerIni(App.Path & "\MARFICE.INI", "BDGENERAL", "PASSW", "")
         
        .Connect = _
        "DSN=" & strBuscar.LeerIni(App.Path & "\MARFICE.INI", "BDGENERAL", "SERVIDOR", "") & ";" & _
        "DSQ=" & strBuscar.LeerIni(App.Path & "\MARFICE.INI", "BDGENERAL", "BDDATOS", "") & ";" & _
        "UID=" & strBuscar.LeerIni(App.Path & "\MARFICE.INI", "BDGENERAL", "USUARIO", "") & ";" & _
        "PWD=" & strBuscar.LeerIni(App.Path & "\MARFICE.INI", "BDGENERAL", "PASSW", "")
        
        .DiscardSavedData = True
        Select Case Left(Trim(Combo1.Text), 2)
            Case "07"
                .formulas(1) = "@titulo ='" & "REPORTE DE NOTAS DE ABONO" & "'"
            Case "08"
                .formulas(1) = "@titulo ='" & "REPORTE DE NOTAS DE CARGO" & "'"
            Case "A"
                .formulas(1) = "@titulo ='" & " REPORTE DE NOTAS DE CARGO/ABONO " & "'"
        End Select
        
        .StoredProcParam(0) = VGCNx.DefaultDatabase
        .StoredProcParam(1) = CStr(DTP_FechaInicio.Value)
        .StoredProcParam(2) = CStr(DTP_FechaFin.Value)
        .StoredProcParam(3) = Trim(adll.ComboDato(Trim(Combo1.Text)))

        If .Status <> 2 Then .Action = 1
        
    End With
    Screen.MousePointer = 1
    Exit Sub
X:
  If Err.Number = 9 Then Resume Next
  Screen.MousePointer = 1
  MsgBox "Error inesperado: " & Err.Number & "  " & Err.Description, vbExclamation
End Sub

