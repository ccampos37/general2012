VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C5FF36B7-A67E-11D3-9D9D-E6F193F7F854}#9.0#0"; "ApliCTxt.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form RptAdel 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reporte de Adelantos de Remuneraciones"
   ClientHeight    =   3045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "RptAdel.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Por fecha de ingreso"
      Height          =   1065
      Left            =   150
      TabIndex        =   8
      Top             =   105
      Visible         =   0   'False
      Width           =   4380
      Begin MSComCtl2.DTPicker xFechaFin 
         Height          =   300
         Left            =   1095
         TabIndex        =   12
         Top             =   660
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   529
         _Version        =   393216
         Format          =   62128129
         CurrentDate     =   36691
      End
      Begin MSComCtl2.DTPicker xFechaIni 
         Height          =   300
         Left            =   1095
         TabIndex        =   11
         Top             =   330
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   529
         _Version        =   393216
         Format          =   62128129
         CurrentDate     =   36691
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         Height          =   195
         Left            =   270
         TabIndex        =   10
         Top             =   713
         Width           =   420
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         Height          =   195
         Left            =   270
         TabIndex        =   9
         Top             =   383
         Width           =   465
      End
   End
   Begin VB.CommandButton CmCancelar 
      Caption         =   "&Cancelar"
      Height          =   390
      Left            =   2483
      TabIndex        =   7
      Top             =   2430
      Width           =   1305
   End
   Begin VB.CommandButton cmImprimir 
      Caption         =   "&Imprimir"
      Height          =   390
      Left            =   893
      TabIndex        =   6
      Top             =   2430
      Width           =   1305
   End
   Begin Crystal.CrystalReport Rpt1 
      Left            =   4065
      Top             =   1320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CheckBox xTodos 
      Caption         =   "Todos los centros de Costo"
      Height          =   300
      Left            =   105
      TabIndex        =   3
      Top             =   1320
      Value           =   1  'Checked
      Width           =   3015
   End
   Begin VB.Frame Frame1 
      Caption         =   "Por mes"
      Height          =   1065
      Left            =   150
      TabIndex        =   0
      Top             =   105
      Width           =   4380
      Begin VB.CheckBox xGeneral 
         Caption         =   "General"
         Height          =   240
         Left            =   1485
         TabIndex        =   4
         Top             =   690
         Visible         =   0   'False
         Width           =   870
      End
      Begin MSComCtl2.DTPicker xMes 
         Height          =   300
         Left            =   1470
         TabIndex        =   1
         Top             =   300
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "MMM/yyyy"
         Format          =   62128131
         CurrentDate     =   36691
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Seleccione mes"
         Height          =   195
         Left            =   195
         TabIndex        =   2
         Top             =   360
         Width           =   1125
      End
   End
   Begin AplisetControlText.Aplitext xCCosto 
      Height          =   285
      Left            =   165
      TabIndex        =   13
      ToolTipText     =   "Centro de Costo"
      Top             =   1920
      Width           =   4305
      _ExtentX        =   7594
      _ExtentY        =   503
      Locked          =   -1  'True
      Text            =   "Todos los Centros de Costos"
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Seleccione el Centro de Costo"
      Height          =   195
      Left            =   135
      TabIndex        =   5
      Top             =   1665
      Width           =   2145
   End
End
Attribute VB_Name = "RptAdel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMCANCELAR_CLICK()
    Unload Me
End Sub

Private Sub CMIMPRIMIR_CLICK()
    Dim CAD As String
    Dim RUTA As String
    If xTodos.Value = 0 Then
        If xCCosto.Tag = "" Then
            MsgBox "FALTA SELECCIONAR CENTRO DE COSTO", vbCritical
            Exit Sub
        End If
    End If
    xMes.Day = 1
    Screen.MousePointer = 11
    Select Case UCase(VPTAREA)
        Case "MES"
            If ExisteTablaAux(" [##RPTADEL01" & VGL_COMPUTER & "] ") Then
                DBSYSTEM.Execute "DROP TABLE  [##RPTADEL01" & VGL_COMPUTER & "] "
            End If
            CAD = "SELECT LTRIM(APEPAT) + ' ' + LTRIM(APEMAT) + ' ' + T.NOMBRE AS NOMBRES, T.CODTRAB, MONTO, CC.NOMBRE, CODCCOSTO INTO  [##RPTADEL01" & VGL_COMPUTER & "]  FROM " & REGSISTEMA.BASESQL & ".dbo.ADEL" & REGSISTEMA.ANNO & " ADEL, " & REGSISTEMA.BASESQL & ".dbo.TRABAJADORES T, " & REGSISTEMA.BASESQL & ".dbo.CCOSTOS CC " & _
            "WHERE CC.CODCCOSTO=T.CCOSTO AND T.CODTRAB=ADEL.CODTRAB AND MES=" & DateSQL(xMes.Value)
            If xTodos.Value = 0 Then
                CAD = CAD + " AND (T.CCOSTO LIKE '" & xCCosto.Tag & "%')"
            End If
            DBSTARPLAN.Execute CAD
            DBSTARPLAN.Execute "EXECUTE [ASISTMP] ' [##RPTADEL01" & VGL_COMPUTER & "] '"
            Rpt1.Connect = "DSN=" & VGL_SERVERREP & ";User=sa;PWD=;DSQ=planilla_pp"
            Rpt1.ReportFileName = REGSISTEMA.REPORTES & "PLAN0033.RPT"
            Rpt1.Formulas(0) = "XEMPRESA='" & REGSISTEMA.EMPRESA & "'"
            Rpt1.Formulas(1) = "XNOMRPT='ADELANTOS DE REMUNERACIONES POR MES'"
            Rpt1.Formulas(2) = "XCONCEPTO='CORRESPONDIENTE A: " & AMESES(Month(xMes.Value)) & " DE " & Year(xMes.Value) & "'"
            Rpt1.StoredProcParam(0) = " [##RPTADEL01" & VGL_COMPUTER & "] "
            Rpt1.Destination = crptToWindow
            Rpt1.WindowState = crptMaximized
            Rpt1.WindowShowRefreshBtn = True
            Rpt1.WindowShowPrintBtn = True
            Rpt1.WindowShowSearchBtn = True
            Rpt1.WindowShowPrintSetupBtn = True
            Rpt1.WindowTitle = "PLAN0033"
            Rpt1.Action = 1
        Case "PENDIENTES"
            If ExisteTablaAux(" [##RPTADEL01" & VGL_COMPUTER & "] ") Then
                DBSYSTEM.Execute "DROP TABLE  [##RPTADEL01" & VGL_COMPUTER & "] "
            End If
            If xGeneral.Value = 0 Then
                CAD = "SELECT LTRIM(APEPAT) + ' ' + LTRIM(APEMAT) + ' ' + T.NOMBRE AS NOMBRES, T.CODTRAB, MONTO, CC.NOMBRE, CODCCOSTO INTO  [##RPTADEL01" & VGL_COMPUTER & "]  FROM " & REGSISTEMA.BASESQL & ".dbo.ADEL" & REGSISTEMA.ANNO & " ADEL, " & REGSISTEMA.BASESQL & ".dbo.TRABAJADORES T, " & REGSISTEMA.BASESQL & ".dbo.CCOSTOS CC " & _
                "WHERE CC.CODCCOSTO=T.CCOSTO AND T.CODTRAB=ADEL.CODTRAB AND NOMBOL=0 AND MES=" & DateSQL(xMes.Value)
            Else
                CAD = "SELECT LTRIM(APEPAT) + ' ' + LTRIM(APEMAT) + ' ' + T.NOMBRE AS NOMBRES, T.CODTRAB, MONTO, CC.NOMBRE, CODCCOSTO INTO  [##RPTADEL01" & VGL_COMPUTER & "]  FROM " & REGSISTEMA.BASESQL & ".dbo.ADEL" & REGSISTEMA.ANNO & " ADEL, " & REGSISTEMA.BASESQL & ".dbo.TRABAJADORES T, " & REGSISTEMA.BASESQL & ".dbo.CCOSTOS CC " & _
                "WHERE CCOSTOS.CODCCOSTO=T.CCOSTO AND T.CODTRAB=ADEL.CODTRAB AND NOMBOL=0"
            End If
            If xTodos.Value = 0 Then
                CAD = CAD + " AND (T.CCOSTO LIKE '" & xCCosto.Tag & "%')"
            End If
            DBSYSTEM.Execute CAD
            DBSTARPLAN.Execute "EXECUTE [ASISTMP] ' [##RPTADEL01" & VGL_COMPUTER & "] '"
            Rpt1.Connect = "DSN=" & VGL_SERVERREP & ";User=sa;PWD=;DSQ=planilla_pp"
            Rpt1.ReportFileName = REGSISTEMA.REPORTES & "PLAN0033.RPT"
            Rpt1.Formulas(0) = "XEMPRESA='" & REGSISTEMA.EMPRESA & "'"
            Rpt1.Formulas(1) = "XNOMRPT='ADELANTOS PENDIENTES POR DESCONTAR'"
            Rpt1.Formulas(2) = "XCONCEPTO='CORRESPONDIENTE A : " & IIf(xGeneral.Value = 0, AMESES(Month(xMes.Value)) & " DE " & Year(xMes.Value), "TODOS LOS MESES") & "'"
            Rpt1.StoredProcParam(0) = " [##RPTADEL01" & VGL_COMPUTER & "] "
            Rpt1.Destination = crptToWindow
            Rpt1.WindowState = crptMaximized
            Rpt1.WindowShowRefreshBtn = True
            Rpt1.WindowShowPrintBtn = True
            Rpt1.WindowShowSearchBtn = True
            Rpt1.WindowShowPrintSetupBtn = True
            Rpt1.WindowTitle = "PLAN0033"
            Rpt1.Action = 1
        Case "INGRESO"
            If ExisteTablaAux(" [##RPTADEL01" & VGL_COMPUTER & "] ") Then
                DBSYSTEM.Execute "DROP TABLE  [##RPTADEL01" & VGL_COMPUTER & "] "
            End If
            CAD = "SELECT LTRIM(APEPAT) + ' ' + LTRIM(APEMAT) + ' ' + T.NOMBRE AS NOMBRES, T.CODTRAB, MONTO, CC.NOMBRE, CODCCOSTO INTO  [##RPTADEL01" & VGL_COMPUTER & "]  FROM " & REGSISTEMA.BASESQL & ".dbo.ADEL" & REGSISTEMA.ANNO & " ADEL, " & REGSISTEMA.BASESQL & ".dbo.TRABAJADORES T, " & REGSISTEMA.BASESQL & ".dbo.CCOSTOS CC " & _
            "WHERE CC.CODCCOSTO=T.CCOSTO AND T.CODTRAB=ADEL.CODTRAB AND (ADEL.FECHAING BETWEEN " & DateSQL(xFechaIni.Value) & " AND " & DateSQL(xFechaFin.Value) & ")"
            If xTodos.Value = 0 Then
                CAD = CAD + " AND (T.CCOSTO LIKE '" & xCCosto.Tag & "%')"
            End If
            DBSYSTEM.Execute CAD
            DBSTARPLAN.Execute "EXECUTE [ASISTMP] ' [##RPTADEL01" & VGL_COMPUTER & "] '"
            Rpt1.Connect = "DSN=" & VGL_SERVERREP & ";User=sa;PWD=;DSQ=planilla_pp"
            Rpt1.ReportFileName = REGSISTEMA.REPORTES & "PLAN0033.RPT"
            Rpt1.Formulas(0) = "XEMPRESA='" & REGSISTEMA.EMPRESA & "'"
            Rpt1.Formulas(1) = "XNOMRPT='ADELANTOS EFECTUADOS POR FECHAS DE INGRESO'"
            Rpt1.Formulas(2) = "XCONCEPTO='CORRESPONDIENTE DESDE : " & xFechaIni.Value & " HASTA " & xFechaFin.Value & "'"
            Rpt1.StoredProcParam(0) = " [##RPTADEL01" & VGL_COMPUTER & "] "
            Rpt1.Destination = crptToWindow
            Rpt1.WindowShowRefreshBtn = True
            Rpt1.WindowShowSearchBtn = True
            Rpt1.WindowShowPrintSetupBtn = True
            Rpt1.WindowShowPrintBtn = True
            Rpt1.WindowState = crptMaximized
            Rpt1.Action = 1
            Rpt1.WindowTitle = "PLAN0033"
    End Select
    Screen.MousePointer = 1
End Sub

Private Sub Form_Load()
    xMes.Value = Date
    xMes.Day = 1
    Select Case VPTAREA
        Case "MES"
            
        Case "PENDIENTES"
            xGeneral.Visible = True
        Case "INGRESO"
            Frame1.Visible = False
            Frame2.Visible = True
    End Select
End Sub

Private Sub XCCOSTO_DBLCLICK()
    Dim RSCCOSTOS As New ADODB.Recordset
    RSCCOSTOS.Open "SELECT CODCCOSTO,NOMBRE FROM CCOSTOS ORDER BY CODCCOSTO", DBSYSTEM, adOpenKeyset, adLockOptimistic
    frmComun.CONECTAR RSCCOSTOS
    frmComun.Show 1
    If VGUTIL(1) <> "" Then
        xCCosto.Text = VGUTIL(1) & " :  " & VGUTIL(2)
        xCCosto.Tag = VGUTIL(1)
        xTodos.Value = 0
    End If
    Set RSCCOSTOS = Nothing
End Sub

Private Sub XGENERAL_Click()
    If xGeneral.Value = 1 Then
        xMes.Enabled = False
    Else
        xMes.Enabled = True
    End If
End Sub

Private Sub XMES_LOSTFOCUS()
    xMes.Day = 1
End Sub

Private Sub XTODOS_Click()
    If xTodos.Value = 1 Then
        xCCosto.Text = "TODOS LOS CENTROS DE COSTOS"
        xCCosto.Tag = ""
    End If
End Sub


