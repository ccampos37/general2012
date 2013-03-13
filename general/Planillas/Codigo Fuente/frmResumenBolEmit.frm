VERSION 5.00
Object = "{C5FF36B7-A67E-11D3-9D9D-E6F193F7F854}#9.0#0"; "ApliCTxt.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmResumenBolEmit 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Resumen de Planillas"
   ClientHeight    =   4875
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4755
   Icon            =   "frmResumenBolEmit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   4755
   StartUpPosition =   3  'Windows Default
   Begin AplisetControlText.Aplitext xConcepto 
      Height          =   240
      Left            =   1020
      TabIndex        =   11
      Top             =   2085
      Width           =   2940
      _ExtentX        =   5186
      _ExtentY        =   423
      Text            =   ""
   End
   Begin AplisetControlText.Aplitext xArea 
      Height          =   240
      Left            =   930
      TabIndex        =   10
      Top             =   3525
      Width           =   3120
      _ExtentX        =   5503
      _ExtentY        =   423
      Locked          =   -1  'True
      Text            =   ""
   End
   Begin Crystal.CrystalReport Reporte 
      Left            =   4110
      Top             =   4215
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmCancelar 
      Caption         =   "&Cancelar"
      Height          =   360
      Left            =   2415
      TabIndex        =   4
      Top             =   4260
      Width           =   1425
   End
   Begin VB.CommandButton cmImprimir 
      Caption         =   "&Imprimir"
      Height          =   360
      Left            =   675
      TabIndex        =   3
      Top             =   4260
      Width           =   1425
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Caption         =   "Tipo de Resumen"
      Height          =   3525
      Left            =   210
      TabIndex        =   0
      Top             =   495
      Width           =   4260
      Begin VB.OptionButton Option3 
         Caption         =   "&Resumen por Concepto de Remuneración"
         Enabled         =   0   'False
         Height          =   195
         Left            =   330
         TabIndex        =   9
         Top             =   2745
         Width           =   3330
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H8000000A&
         Caption         =   "Filtro por Centro de Costo"
         Enabled         =   0   'False
         Height          =   1275
         Left            =   585
         TabIndex        =   5
         Top             =   1290
         Width           =   3300
         Begin VB.ComboBox xNivel 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmResumenBolEmit.frx":000C
            Left            =   1650
            List            =   "frmResumenBolEmit.frx":001F
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   780
            Width           =   1545
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            Caption         =   "Nivel de Impresión:"
            Height          =   210
            Left            =   135
            TabIndex        =   8
            Top             =   825
            Width           =   1350
         End
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H8000000A&
         Caption         =   "Resumen por  Centro de Costo"
         Height          =   195
         Left            =   330
         TabIndex        =   2
         Top             =   885
         Width           =   2490
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H8000000A&
         Caption         =   "Resumen General"
         Height          =   195
         Left            =   330
         TabIndex        =   1
         Top             =   540
         Value           =   -1  'True
         Width           =   2085
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   3465
         Top             =   330
         Width           =   480
      End
   End
   Begin VB.Label xPlanilla 
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   210
      TabIndex        =   7
      Top             =   120
      Width           =   4260
   End
End
Attribute VB_Name = "frmResumenBolEmit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMCANCELAR_CLICK()
    Unload Me
End Sub

Private Sub CMIMPRIMIR_CLICK()
On Error GoTo Err1
    
    Screen.MousePointer = 11
    Dim xMes As String
    Dim RSAUX As New ADODB.Recordset, RSSUM As New ADODB.Recordset
    Dim K As Integer
    Dim SNOMBOL As String
    
    Dim FMES As Date
    
    SNOMBOL = Right(frBolEmit.Lista.SelectedItem.KEY, Len(frBolEmit.Lista.SelectedItem.KEY) - 1)
    FMES = DevuelveValor("SELECT MES FROM NOMBOL WHERE CODIGO=" & SNOMBOL, DBSYSTEM)
    xMes = Format(Month(FMES), "00") & Year(FMES)
    
    If Option1.Value Then
        Dim VALOR1, VALOR2, VALOR3  As Integer
        CambiaPanelBD True
         VALOR1 = DevuelveValor("SELECT SUM(MONTO) AS T1 FROM PAGOSCTA WHERE TIPO=1 AND CODNOMBOL=" & SNOMBOL & " AND CODTRAB IN (SELECT CODTRAB FROM  [##_TMPLSTBOL" & VGL_COMPUTER & "]  WHERE NOMBOL=" & SNOMBOL & " AND TIPOBOLETA='B'  )", DBSYSTEM)
         VALOR2 = DevuelveValor("SELECT SUM(MONTO) AS T1 FROM ADEL2000 WHERE ORIGEN=" & SNOMBOL & " AND CODTRAB IN (SELECT CODTRAB FROM  [##_TMPLSTBOL" & VGL_COMPUTER & "]  WHERE NOMBOL=" & SNOMBOL & ")", DBSYSTEM)
         VALOR3 = DevuelveValor("SELECT SUM(MONTO) AS T1 FROM PAGOSCTA WHERE TIPO=2 AND CODNOMBOL=" & SNOMBOL & " AND CODTRAB IN (SELECT CODTRAB FROM  [##_TMPLSTBOL" & VGL_COMPUTER & "]  WHERE NOMBOL=" & SNOMBOL & " AND TIPOBOLETA='B' )", DBSYSTEM)
         DBSTARPLAN.Execute "EXECUTE CURSOR_RESUMEN '" & REGSISTEMA.BASESQL & "' ,'" & xMes & "', " & VALOR1 & " ," & VALOR2 & " ," & VALOR3 & ",'" & VGL_COMPUTER & "'"
        If ExisteTablaAux("[##TMPRESUMENGENERAL_AUX" & VGL_COMPUTER & "]") Then
            DBSYSTEM.Execute "DROP TABLE [##TMPRESUMENGENERAL_AUX" & VGL_COMPUTER & "]"
        End If
        DBSYSTEM.Execute "SELECT * INTO [##TMPRESUMENGENERAL_AUX" & VGL_COMPUTER & "] FROM [##TMPRESUMENGENERAL" & VGL_COMPUTER & "]"
        DBSYSTEM.Execute "DELETE FROM [##TMPRESUMENGENERAL_AUX" & VGL_COMPUTER & "]"
        
        For K = 1 To frBolEmit.Lista.ListItems.Count
            If frBolEmit.Lista.ListItems(K).Checked = True Then
                SNOMBOL = Right(frBolEmit.Lista.ListItems(K).KEY, Len(frBolEmit.Lista.ListItems(K).KEY) - 1)
                FMES = DevuelveValor("SELECT MES FROM NOMBOL WHERE CODIGO=" & SNOMBOL, DBSYSTEM)
                xMes = Format(Month(FMES), "00") & Year(FMES)
                VALOR1 = DevuelveValor("SELECT SUM(MONTO) AS T1 FROM PAGOSCTA WHERE TIPO=1 AND CODNOMBOL=" & SNOMBOL & " AND CODTRAB IN (SELECT CODTRAB FROM  [##_TMPLSTBOL" & VGL_COMPUTER & "]  WHERE NOMBOL=" & SNOMBOL & " AND TIPOBOLETA='B' )", DBSYSTEM)
                VALOR2 = DevuelveValor("SELECT SUM(MONTO) AS T1 FROM ADEL2000 WHERE ORIGEN=" & SNOMBOL & " AND CODTRAB IN (SELECT CODTRAB FROM  [##_TMPLSTBOL" & VGL_COMPUTER & "]  WHERE NOMBOL=" & SNOMBOL & ")", DBSYSTEM)
                VALOR3 = DevuelveValor("SELECT SUM(MONTO) AS T1 FROM PAGOSCTA WHERE TIPO=2 AND CODNOMBOL=" & SNOMBOL & " AND CODTRAB IN (SELECT CODTRAB FROM  [##_TMPLSTBOL" & VGL_COMPUTER & "]  WHERE NOMBOL=" & SNOMBOL & " AND TIPOBOLETA='B' )", DBSYSTEM)
                DBSTARPLAN.Execute "EXECUTE CURSOR_RESUMEN '" & REGSISTEMA.BASESQL & "' ,'" & xMes & "', " & VALOR1 & " ," & VALOR2 & " ," & VALOR3 & ",'" & VGL_COMPUTER & "'"
                Set RSSUM = Nothing
                RSSUM.Open "SELECT SUM(XREDONDEO) AS REDONDEO FROM BOL" & xMes & " MOV WHERE INUMBOL IN (SELECT INUMBOL FROM  [##_TMPLSTBOL" & VGL_COMPUTER & "] )", DBSYSTEM, adOpenStatic, adLockReadOnly
                If Not (RSSUM.RecordCount = 0 Or RSSUM.EOF) Then DBSTARPLAN.Execute "INSERT INTO [" & "##TMPRESUMENGENERAL" & VGL_COMPUTER & "] VALUES ('REDONDEO','REDONDEO'," & Round(IIf(IsNull(RSSUM!REDONDEO), 0, RSSUM!REDONDEO), 3) & ",4)"
                DBSYSTEM.Execute "INSERT INTO [##TMPRESUMENGENERAL_AUX" & VGL_COMPUTER & "] SELECT * FROM [##TMPRESUMENGENERAL" & VGL_COMPUTER & "]"
            End If
        Next K
            
            Set RSAUX = Nothing
            DBSTARPLAN.Execute "EXECUTE [ASISTMP] '[##TMPRESUMENGENERAL_AUX" & VGL_COMPUTER & "]'"
        With Reporte
            .Reset
            .WindowTitle = "RESUMEN GENERAL DE PLANILLAS"
            .ReportFileName = REGSISTEMA.REPORTES & "PLAN0042.RPT"
            .Connect = "DSN=" & VGL_SERVERREP & ";User=sa;PWD=;DSQ=" & VGL_BASE & ""
            .StoredProcParam(0) = "[##TMPRESUMENGENERAL_AUX" & VGL_COMPUTER & "]"
            .Destination = crptToWindow
            .WindowState = crptMaximized
            .Formulas(0) = "XEMPRESA='" & REGSISTEMA.EMPRESA & "'"
            .Formulas(1) = "XMES='" & frBolEmit.Lista.SelectedItem.Text & "'"
            .Formulas(2) = "XHORA='" & Time & "'"
            .WindowShowPrintBtn = True
            .WindowShowPrintSetupBtn = True
            .WindowShowRefreshBtn = True
            .WindowShowSearchBtn = True
            Screen.MousePointer = 1
            CambiaPanelBD False
            If Reporte.Status <> 2 Then .Action = 1
        End With
    End If
    
    If Option2.Value Then
        CambiaPanelBD True
        Dim SALL As String
        SALL = ""
        If xArea.Text <> "" Then SALL = " AND CCOSTO=" & Trim(xArea.Tag) & " "
        DBSTARPLAN.Execute "EXECUTE CURSOR_PLANILLA_NIVELES '" & REGSISTEMA.BASESQL & "', '" & SALL & "', '" & xMes & "','" & VGL_COMPUTER & "'"
        Call CreaTempCostos(DBSYSTEM, "CCOSTOS", "CODCCOSTO")
        DBSTARPLAN.Execute "EXECUTE CURSOR_PLANILLA_NIVELES_2 '" & REGSISTEMA.BASESQL & "','" & VGL_COMPUTER & "'"
        
        'COMIENZA LA MODIFICACION PARA EL PROCESO DE LISTADO ENTRE MESES
        
        If ExisteTablaAux("[##TMPRESUMENCCOSTO_AUX" & VGL_COMPUTER & "]") Then DBSYSTEM.Execute "DROP TABLE [##TMPRESUMENCCOSTO_AUX" & VGL_COMPUTER & "]"
        DBSYSTEM.Execute "SELECT * INTO  [##TMPRESUMENCCOSTO_AUX" & VGL_COMPUTER & "] FROM [##TMPRESUMENCCOSTO" & VGL_COMPUTER & "]"
        DBSYSTEM.Execute "DELETE FROM [##TMPRESUMENCCOSTO_AUX" & VGL_COMPUTER & "]"
        
        For K = 1 To frBolEmit.Lista.ListItems.Count
            If frBolEmit.Lista.ListItems(K).Checked = True Then
                SNOMBOL = Right(frBolEmit.Lista.ListItems(K).KEY, Len(frBolEmit.Lista.ListItems(K).KEY) - 1)
                FMES = DevuelveValor("SELECT MES FROM NOMBOL WHERE CODIGO=" & SNOMBOL, DBSYSTEM)
                xMes = Format(Month(FMES), "00") & Year(FMES)
        
                SALL = ""
                If xArea.Text <> "" Then SALL = " AND CCOSTO=" & Trim(xArea.Tag) & " "
                DBSTARPLAN.Execute "EXECUTE CURSOR_PLANILLA_NIVELES '" & REGSISTEMA.BASESQL & "', '" & SALL & "', '" & xMes & "','" & VGL_COMPUTER & "'"
                'CREAR TABLA TEMPORAL CON LOS NIVELES
                Call CreaTempCostos(DBSYSTEM, "CCOSTOS", "CODCCOSTO")
                'crear duplicados para almacenar
                DBSTARPLAN.Execute "EXECUTE CURSOR_PLANILLA_NIVELES_2 '" & REGSISTEMA.BASESQL & "','" & VGL_COMPUTER & "'"
                'AHORA EL NUEVO TMP LO TIENE
                DBSYSTEM.Execute "INSERT  [##TMPRESUMENCCOSTO_AUX" & VGL_COMPUTER & "] SELECT * FROM [##TMPRESUMENCCOSTO" & VGL_COMPUTER & "]"
            End If
        Next K
        DBSYSTEM.Execute "DELETE FROM [##TMPRESUMENCCOSTO" & VGL_COMPUTER & "]"
        DBSYSTEM.Execute "INSERT INTO [##TMPRESUMENCCOSTO" & VGL_COMPUTER & "]  SELECT * FROM [##TMPRESUMENCCOSTO_AUX" & VGL_COMPUTER & "]"
        DBSYSTEM.Execute "" & _
        " Insert into [##TMPRESUMENCCOSTO" & VGL_COMPUTER & "] ( " & _
        " CODIGO,NOMBRE,MONTOVAL,TIPO,CCOSTO,CODCCOSTO,NIVEL1) " & _
        " SELECT CODIGO='_Adelanto',NOMBRE='_Adelanto de Quincena',MONTOVAL=SUM(ADEL.MONTO),TIPO=2, CCOSTO=LEFT(b.CCOSTO,2), " & _
        " CODCCOSTO=LEFT(b.CCOSTO,2),NIVEL1=LEFT(b.CCOSTO,2) " & _
        " FROM   DBO.ADEL2000 ADEL,DBO.TRABAJADORES B " & _
        "  WHERE ADEL.CODTRAB=B.CODTRAB AND  ADEL.CODTRAB in (SELECT codtrab FROM ##_TMPLSTBOL" & VGL_COMPUTER & ") and   " & _
        "  ORIGEN=" & SNOMBOL & _
        "   GROUP BY LEFT(b.CCOSTO,2) " & _
        "  Union All " & _
        "  SELECT CODIGO=CASE WHEN ADEL.TIPO=1 THEN '_OtrIngCta' else '_OtrEgrCta' end, " & _
        "  NOMBRE=case when ADEL.TIPO=1 then '_Otros Ingresos Cta' else '_Otros Egresos Cta' end, " & _
        "  MONTOVAL=SUM(ADEL.MONTO),TIPO=ADEL.TIPO, " & _
        "  CCOSTO=LEFT(b.CCOSTO,2), " & _
        "  CODCCOSTO=LEFT(b.CCOSTO,2),NIVEL1=LEFT(b.CCOSTO,2) " & _
        "  FROM   DBO.PAGOSCTA ADEL,DBO.TRABAJADORES B " & _
        "  WHERE ADEL.CODTRAB=B.CODTRAB and ADEL.CODTRAB in (SELECT codtrab FROM ##_TMPLSTBOL" & VGL_COMPUTER & ")  and " & _
        "  ADEL.TIPOBOLETA='B' AND ADEL.CODNOMBOL=" & SNOMBOL & _
        "  GROUP BY ADEL.TIPO,LEFT(b.CCOSTO,2) "
       
        
        Dim NOMREP As String
        'Set RSAUX = Nothing
        With Reporte
            .Reset
            Select Case xNivel.ListIndex
                Case 0
                    DBSTARPLAN.Execute "EXECUTE [ASISTMPBOL] '[##TMPRESUMENCCOSTO" & VGL_COMPUTER & "]','[##TMPCCO" & VGL_COMPUTER & "]','NIVEL1','CODCCOSTO'"
                    .ReportFileName = REGSISTEMA.REPORTES & "PNV10047.RPT"
                    NOMREP = "PNV10047.RPT 1ER NIVEL "
                Case 1
                    DBSTARPLAN.Execute "EXECUTE [ASISTMPBOL] '[##TMPRESUMENCCOSTO" & VGL_COMPUTER & "]','[##TMPCCO" & VGL_COMPUTER & "]','NIVEL2','CODCCOSTO'"
                    .ReportFileName = REGSISTEMA.REPORTES & "PNV20047.RPT"
                    NOMREP = "PNV20047.RPT 2DO NIVEL "
                Case 2
                    DBSTARPLAN.Execute "EXECUTE [ASISTMPBOL] '[##TMPRESUMENCCOSTO" & VGL_COMPUTER & "]' ,'[##TMPCCO" & VGL_COMPUTER & "]','NIVEL3','CODCCOSTO'"
                    .ReportFileName = REGSISTEMA.REPORTES & "PNV30047.RPT"
                    NOMREP = "PNV30047.RPT 3ER NIVEL "
                    
                Case 3
                    DBSTARPLAN.Execute "EXECUTE [ASISTMPBOL] '[##TMPRESUMENCCOSTO" & VGL_COMPUTER & "]','[##TMPCCO" & VGL_COMPUTER & "]','NIVEL4','CODCCOSTO'"
                    .ReportFileName = REGSISTEMA.REPORTES & "PNV40047.RPT"
                    NOMREP = "PNV40047.RPT 4TO NIVEL "
                Case 4
                    DBSTARPLAN.Execute "EXECUTE [ASISTMPBOL] '[##TMPRESUMENCCOSTO" & VGL_COMPUTER & "]','[##TMPCCO" & VGL_COMPUTER & "]','NIVEL5','CODCCOSTO'"
                    .ReportFileName = REGSISTEMA.REPORTES & "PNV50047.RPT"
                    NOMREP = "PNV50047.RPT 5TO NIVEL "
            End Select
            .WindowTitle = NOMREP & "- RESUMEN GENERAL DE PLANILLAS POR CENTRO DE COSTO"
            .Connect = "DSN=" & VGL_SERVERREP & ";User=sa;PWD=;DSQ=planilla_pp"
            .StoredProcParam(0) = "[##TMPRESUMENCCOSTO" & VGL_COMPUTER & "]"
            .StoredProcParam(1) = "[##TMPCCO" & VGL_COMPUTER & "]"
            .StoredProcParam(2) = "NIVEL" & CStr(xNivel.ListIndex + 1)
            .StoredProcParam(3) = "CODCCOSTO"
            .Destination = crptToWindow
            .WindowState = crptMaximized
            .Formulas(0) = "XEMPRESA='" & REGSISTEMA.EMPRESA & "'"
            .Formulas(1) = "XMES='" & frBolEmit.Lista.SelectedItem.Text & "'"
            .Formulas(2) = "XHORA='" & Time & "'"
            .WindowShowPrintBtn = True
            .WindowShowPrintSetupBtn = True
            .WindowShowRefreshBtn = True
            .WindowShowSearchBtn = True
            If Reporte.Status <> 2 Then .Action = 1
            Screen.MousePointer = 1
            CambiaPanelBD False
        End With
    End If
    If Option3.Value Then
        CambiaPanelBD True
        Screen.MousePointer = 11
        'COMIENZA LA MODIFICACION PARA EL PROCESO DE LISTADO ENTRE MESES
        DBSTARPLAN.Execute "SP_CON_REMUNE '" & REGSISTEMA.BASESQL & "', '" & xMes & "', '" & xConcepto.Tag & "','" & VGL_COMPUTER & "'"
        If ExisteTablaAux("[##TMPRESUMENCNPT_AUX" & VGL_COMPUTER & "]") Then DBSYSTEM.Execute "DROP TABLE [##TMPRESUMENCNPT_AUX" & VGL_COMPUTER & "]"
        DBSYSTEM.Execute "SELECT * INTO [##TMPRESUMENCNPT_AUX" & VGL_COMPUTER & "] FROM [##TMPRESUMENCNPT" & VGL_COMPUTER & "]"
        DBSYSTEM.Execute "DELETE FROM [##TMPRESUMENCNPT_AUX" & VGL_COMPUTER & "]"
        
        For K = 1 To frBolEmit.Lista.ListItems.Count
            If frBolEmit.Lista.ListItems(K).Checked = True Then
                SNOMBOL = Right(frBolEmit.Lista.ListItems(K).KEY, Len(frBolEmit.Lista.ListItems(K).KEY) - 1)
                FMES = DevuelveValor("SELECT MES FROM NOMBOL WHERE CODIGO=" & SNOMBOL, DBSYSTEM)
                xMes = Format(Month(FMES), "00") & Year(FMES)
                If Len(Trim(xConcepto.Tag)) = 0 Then
                    xConcepto.Tag = "%"
                End If
                DBSTARPLAN.Execute "SP_CON_REMUNE '" & REGSISTEMA.BASESQL & "', '" & xMes & "', '" & xConcepto.Tag & "','" & VGL_COMPUTER & "'"
                'AHORA EL NUEVO TMP LO TIENE
                DBSYSTEM.Execute "INSERT INTO [##TMPRESUMENCNPT_AUX" & VGL_COMPUTER & "] SELECT * FROM [" & "##TMPRESUMENCNPT" & VGL_COMPUTER & "]"
             End If
         Next K
         DBSYSTEM.Execute "DELETE FROM [##TMPRESUMENCNPT" & VGL_COMPUTER & "]"
         DBSYSTEM.Execute "INSERT INTO [##TMPRESUMENCNPT" & VGL_COMPUTER & "]  SELECT * FROM [##TMPRESUMENCNPT_AUX" & VGL_COMPUTER & "]"
'        If ExisteTablaAux("TMPRESUMENCNPT") Then DBSYSTEM.Execute "DROP TABLE TMPRESUMENCNPT"
'        If ExisteTablaAux("TMPCNPT") Then DBSYSTEM.Execute "DROP TABLE TMPCNPT"
'        If ExisteTablaAux("MOV") Then DBSYSTEM.Execute "DROP TABLE MOV"
'        DBSYSTEM.Execute "SELECT * INTO MOV FROM MOV" & xMes & " IN '" & REGSISTEMA.PATHEMPRESA & "\PLANILLA.MDB'"
'        DBSYSTEM.Execute "SELECT * INTO TMPCNPT FROM CONCEPTOS IN '" & REGSISTEMA.PATHEMPRESA & "\PLANILLA.MDB'"
'        DBSYSTEM.Execute "SELECT CODTRAB, NOMBRES, TMPCNPT.NOMBRE, PERIODO, MONTO INTO TMPRESUMENCNPT FROM _TMPLSTBOL LISTA, TMPCNPT, MOV WHERE LISTA.INUMBOL=MOV.INUMBOL AND MOV.CONCEPTO=TMPCNPT.CODIGO AND TMPCNPT.CODIGO='" & xConcepto.Tag & "'"
        With Reporte
        'MODIFICADO POR BASILIO
            .Reset
            .WindowTitle = "RESUMEN POR CONCEPTO DE REMUNERACIÓN"
            .Connect = "DSN=" & VGL_SERVER & ";User=sa;PWD=;DSQ=" & REGSISTEMA.BASESQL & ""
            .ReportFileName = REGSISTEMA.REPORTES & "pruebaX.RPT" '"PLAN0046.RPT"
            .StoredProcParam(0) = "[##TMPRESUMENCNPT" & VGL_COMPUTER & "]" 'REGSISTEMA.BASESQL
            '.StoredProcParam(1) = xMes
            '.StoredProcParam(2) = xConcepto.Tag
            '.StoredProcParam(3) = VGL_COMPUTER
            .Destination = crptToWindow
            .WindowState = crptMaximized
            .Formulas(0) = "XEMPRESA='" & REGSISTEMA.EMPRESA & "'"
            .Formulas(1) = "XMES='" & frBolEmit.Lista.SelectedItem.Text & "'"
            .Formulas(2) = "XRUC='RUC: " & REGSISTEMA.RUC & "'"
            .Formulas(3) = "XDIRECCION='" & DevuelveValor("SELECT DIRECCIÓN FROM EMPRESA", DBSYSTEM) & "'"
            .WindowShowPrintBtn = True
            .WindowShowPrintSetupBtn = True
            .WindowShowRefreshBtn = True
            .WindowShowSearchBtn = True
            Screen.MousePointer = 1
            CambiaPanelBD False
            If Reporte.Status <> 2 Then .Action = 1
        End With
    End If
    Exit Sub
Err1:
    MsgBox ERR.Description
    Resume Next
End Sub


Private Sub Form_Load()
If frBolEmit.xVistaMes.List(1) = "Por Rango de Meses" Then
    xPlanilla.Caption = Format(frBolEmit.xFechaIni, "mmmm yyyy") & " - " & Format(frBolEmit.xFechaFin, "mmmm yyyy")
Else
    xPlanilla.Caption = frBolEmit.Lista.SelectedItem.Text
End If
End Sub

Private Sub OPTION1_CLICK()
    Frame2.Enabled = False
End Sub

Private Sub OPTION2_Click()
    Frame2.Enabled = True
    xNivel.ListIndex = 0
End Sub

Private Sub OPTION3_Click()
    Frame2.Enabled = False
End Sub

Private Sub XAREA_Click()
    Dim RSAUX As ADODB.Recordset
    Set RSAUX = New ADODB.Recordset
    RSAUX.Open "SELECT CODCCOSTO, NOMBRE FROM CCOSTOS ORDER BY CODCCOSTO", DBSYSTEM, adOpenStatic
    If RSAUX.RecordCount = 0 Then
        MsgBox "No se han encontrado registros de Centro de Costo", vbCritical
        Set RSAUX = Nothing
        Exit Sub
    End If
    frmComun.CONECTAR RSAUX
    frmComun.Show 1
    If VGUTIL(1) <> "" Then
        xArea.Text = RSAUX!CODCCOSTO & " - " & RSAUX!NOMBRE
        xArea.Tag = RSAUX!CODCCOSTO
    End If
    Set RSAUX = Nothing
End Sub

Private Sub XAREA_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then
        xArea.Tag = "": xArea.Text = ""
    End If
End Sub

Private Sub XCONCEPTO_DBLCLICK()
    Dim RSAUX As New ADODB.Recordset
    RSAUX.Open "CONCEPTOS", DBSYSTEM, adOpenStatic, adLockReadOnly
    If RSAUX.EOF Or RSAUX.RecordCount = 0 Then
        MsgBox "No existen conceptos de remuneraciones", vbInformation
        Set RSAUX = Nothing
        Exit Sub
    End If
    frmComun.CONECTAR RSAUX
    frmComun.Show 1
    If VGUTIL(1) <> "" Then
        xConcepto.Text = VGUTIL(2)
        xConcepto.Tag = VGUTIL(1)
    End If
    Set RSAUX = Nothing
End Sub


