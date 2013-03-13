VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C5FF36B7-A67E-11D3-9D9D-E6F193F7F854}#9.0#0"; "ApliCTxt.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frRngFch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingresar Rango de Fecha"
   ClientHeight    =   4470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3960
   Icon            =   "frRngFch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   3960
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Otros filtros"
      Height          =   1680
      Left            =   157
      TabIndex        =   7
      Top             =   2085
      Width           =   3690
      Begin AplisetControlText.Aplitext xCampo 
         Height          =   285
         Left            =   165
         TabIndex        =   11
         Top             =   1230
         Width           =   3285
         _ExtentX        =   5794
         _ExtentY        =   503
         Locked          =   -1  'True
         Text            =   ""
      End
      Begin AplisetControlText.Aplitext xCCosto 
         Height          =   285
         Left            =   165
         TabIndex        =   9
         Top             =   570
         Width           =   3285
         _ExtentX        =   5794
         _ExtentY        =   503
         Locked          =   -1  'True
         Text            =   ""
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Campo a Imprimir"
         Height          =   195
         Left            =   165
         TabIndex        =   10
         Top             =   990
         Width           =   1200
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Area de Trabajo"
         Height          =   195
         Left            =   180
         TabIndex        =   8
         Top             =   360
         Width           =   1140
      End
   End
   Begin Crystal.CrystalReport Rpt01 
      Left            =   6315
      Top             =   6465
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmCancelar 
      Caption         =   "&Cancelar"
      Height          =   390
      Left            =   2115
      TabIndex        =   6
      Top             =   3930
      Width           =   1095
   End
   Begin VB.CommandButton cmAceptar 
      Caption         =   "&Aceptar"
      Height          =   390
      Left            =   795
      TabIndex        =   5
      Top             =   3930
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Rango de Fechas para el Reporte"
      Height          =   1890
      Left            =   157
      TabIndex        =   0
      Top             =   120
      Width           =   3690
      Begin MSComCtl2.DTPicker xFechaFin 
         Height          =   315
         Left            =   1230
         TabIndex        =   4
         Top             =   1365
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         Format          =   61865985
         CurrentDate     =   36689
      End
      Begin MSComCtl2.DTPicker xFechaIni 
         Height          =   315
         Left            =   1230
         TabIndex        =   3
         Top             =   675
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         Format          =   61865985
         CurrentDate     =   36689
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   255
         Picture         =   "frRngFch.frx":0442
         Top             =   465
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Final"
         Height          =   195
         Left            =   1230
         TabIndex        =   2
         Top             =   1140
         Width           =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Inicial"
         Height          =   195
         Left            =   1230
         TabIndex        =   1
         Top             =   435
         Width           =   900
      End
   End
End
Attribute VB_Name = "frRngFch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMACEPTAR_CLICK()
Dim TABLA_TMPORAL As String, OPCION As Integer, SQL  As String
    Screen.MousePointer = 11
    Rpt01.Reset
    Select Case VPTAREA
        Case "PLAN0003.RPT"
            TABLA_TMPORAL = " [##RPTACTUAL" & VGL_COMPUTER & "] "
            If ExisteTablaAux(" [##RPTACTUAL" & VGL_COMPUTER & "] ") Then
                DBSYSTEM.Execute "DROP TABLE  [##RPTACTUAL" & VGL_COMPUTER & "] "
            End If
            SQL = "SELECT NOMBRES, DAY(DIA) AS DIAT, VALOR INTO  [##RPTACTUAL" & VGL_COMPUTER & "]  FROM " & REGSISTEMA.BASESQL & ".dbo.VWTRABAJ W, " & REGSISTEMA.BASESQL & ".dbo.ASIS" & REGSISTEMA.ANNO & " A " & _
            "WHERE A.CODTRAB=W.CODTRAB AND (DIA BETWEEN " & DateSQL(xFechaIni.Value) & " AND " & DateSQL(xFechaFin.Value) & ") AND  W.CODAREA='" & xCCosto.Tag & "' AND A.CONCEPTO='" & xCampo.Tag & "'"
            DBSYSTEM.Execute SQL
        Case "PLAN0004.RPT"
            TABLA_TMPORAL = " [##RPTASIS002" & VGL_COMPUTER & "] "
            If ExisteTablaAux(" [##RPTASIS002" & VGL_COMPUTER & "] ") Then
                DBSYSTEM.Execute "DROP TABLE  [##RPTASIS002" & VGL_COMPUTER & "] "
            End If
            SQL = "SELECT DIA, C.NOMBRE, VALOR INTO  [##RPTASIS002" & VGL_COMPUTER & "]  FROM " & REGSISTEMA.BASESQL & ".dbo.CONCEPTOS C, " & REGSISTEMA.BASESQL & ".dbo.ASIS" & REGSISTEMA.ANNO & " A  " & _
            " WHERE A.CONCEPTO=C.CODIGO AND (DIA BETWEEN " & DateSQL(xFechaIni.Value) & " AND " & DateSQL(xFechaFin.Value) & ") AND A.CODTRAB='" & xCCosto.Tag & "'"
            DBSYSTEM.Execute SQL
        Case "PLAN0005.RPT"
            TABLA_TMPORAL = " [##RPTASIS003" & VGL_COMPUTER & "] "
            If ExisteTablaAux(" [##RPTASIS003" & VGL_COMPUTER & "] ") Then
                DBSYSTEM.Execute "DROP TABLE  [##RPTASIS003" & VGL_COMPUTER & "] "
            End If
            SQL = "SELECT NOMBRES, C.NOMBRE, VALOR INTO   [##RPTASIS003" & VGL_COMPUTER & "]  FROM " & REGSISTEMA.BASESQL & ".dbo.VWTRABAJ W, " & REGSISTEMA.BASESQL & ".dbo.CONCEPTOS C, " & REGSISTEMA.BASESQL & ".dbo.ASIS" & REGSISTEMA.ANNO & " A " & _
            " WHERE A.CONCEPTO=C.CODIGO  AND A.CODTRAB=W.CODTRAB AND (DIA BETWEEN " & DateSQL(xFechaIni.Value) & " AND " & DateSQL(xFechaFin.Value) & ") AND  W.CODAREA='" & xCCosto.Tag & "'"
            DBSYSTEM.Execute SQL
        Case Else
            MsgBox "REPORTE NO DISPONIBLE", vbCritical
            Screen.MousePointer = 1
            Exit Sub
    End Select
    DBSTARPLAN.Execute "EXECUTE [ASISTMP] '" & TABLA_TMPORAL & "'"
    Rpt01.ReportFileName = REGSISTEMA.REPORTES & VPTAREA
    
    'Rpt01.LogOnServer "pdssql.dll", VGL_SERVERREP, "MARFICE_PP", "SOPORTE", "SOPORTE"
    Rpt01.Connect = "DSN=" & VGL_SERVERREP & ";User=sa;PWD=;DSQ=planilla_pp"
    Rpt01.StoredProcParam(0) = TABLA_TMPORAL
    Rpt01.Destination = crptToWindow
    Rpt01.WindowState = crptMaximized
    Rpt01.WindowShowPrintBtn = True
    Rpt01.WindowShowRefreshBtn = True
    Rpt01.WindowShowSearchBtn = True
    Rpt01.WindowShowPrintSetupBtn = True

    Rpt01.WindowTitle = VPTAREA
    Rpt01.Formulas(0) = "XEMPRESA='" & REGSISTEMA.EMPRESA & "'"
    Rpt01.Formulas(1) = "XCAMPO='" & xCCosto.Text & "'"
    Rpt01.Formulas(2) = "XMES='DESDE: " & xFechaIni.Value & " HASTA: " & xFechaFin.Value & "'"
    If Rpt01.Status <> 2 Then Rpt01.Action = 1
    Screen.MousePointer = 1
End Sub

Private Sub CMCANCELAR_CLICK()
    Unload Me
End Sub

Private Sub XCAMPO_DblClick()
    Dim RSCCOSTOS As New ADODB.Recordset
    RSCCOSTOS.Open "SELECT CODIGO, NOMBRE FROM CONCEPTOS ORDER BY NOMBRE", DBSYSTEM, adOpenKeyset, adLockOptimistic
    frmComun.CONECTAR RSCCOSTOS
    frmComun.Show 1
    If VGUTIL(1) <> "" Then
        xCampo.Text = VGUTIL(1) & " :  " & VGUTIL(2)
        xCampo.Tag = VGUTIL(1)
    End If
    Set RSCCOSTOS = Nothing
End Sub

Private Sub XCCOSTO_DBLCLICK()
    Dim RSCCOSTOS As New ADODB.Recordset
    If UCase(Label3.Caption) = "TRABAJADOR" Then
        RSCCOSTOS.Open "SELECT CODTRAB, NOMBRES FROM VWTRABAJ", DBSYSTEM, adOpenStatic
    Else
        RSCCOSTOS.Open "SELECT CODCCOSTO,NOMBRE FROM AREASTRAB ORDER BY CODCCOSTO", DBSYSTEM, adOpenKeyset, adLockOptimistic
    End If
    frmComun.CONECTAR RSCCOSTOS
    frmComun.Show 1
    If VGUTIL(1) <> "" Then
        xCCosto.Text = VGUTIL(1) & " :  " & VGUTIL(2)
        xCCosto.Tag = VGUTIL(1)
    End If
    Set RSCCOSTOS = Nothing
End Sub


