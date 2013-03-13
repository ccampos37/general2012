VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frAdminCTS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Depósitos de C.T.S."
   ClientHeight    =   5805
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8610
   Icon            =   "frAdminCTS.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   8610
   Begin VB.CommandButton cmCalcMes 
      Caption         =   "Cálculo Mensual"
      Height          =   465
      Left            =   1425
      TabIndex        =   14
      Top             =   4470
      Width           =   1140
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Certificado de CTS Mensual"
      Height          =   465
      Left            =   2625
      TabIndex        =   13
      Top             =   4485
      Width           =   1140
   End
   Begin Crystal.CrystalReport Reporte 
      Left            =   3555
      Top             =   2265
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Planilla de C.&T.S."
      Height          =   465
      Left            =   210
      TabIndex        =   9
      Top             =   4485
      Width           =   1155
   End
   Begin VB.CommandButton cmConsulta 
      Caption         =   "&Consulta"
      Height          =   405
      Left            =   7230
      TabIndex        =   2
      Top             =   3150
      Width           =   1275
   End
   Begin VB.CommandButton cmCustodia 
      Caption         =   "P&oner en Custodia"
      Height          =   465
      Left            =   4590
      TabIndex        =   10
      Top             =   4485
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.CommandButton cmListado 
      Caption         =   "&Listado al Banco"
      Height          =   465
      Left            =   5805
      TabIndex        =   11
      Top             =   4485
      Width           =   1155
   End
   Begin VB.CommandButton cmCerrar 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   405
      Left            =   7230
      TabIndex        =   6
      Top             =   5310
      Width           =   1275
   End
   Begin VB.CommandButton cmAceptar 
      Caption         =   "&Aceptar"
      Height          =   405
      Left            =   7230
      TabIndex        =   5
      Top             =   4770
      Width           =   1275
   End
   Begin VB.CommandButton cmEliminar 
      Caption         =   "&Eliminar"
      Height          =   405
      Left            =   7230
      TabIndex        =   3
      Top             =   3690
      Width           =   1275
   End
   Begin VB.CommandButton cmModificar 
      Caption         =   "&Modificar"
      Height          =   405
      Left            =   7230
      TabIndex        =   4
      Top             =   4230
      Width           =   1275
   End
   Begin VB.CommandButton cmPrueba 
      Caption         =   "&Prueba Cálculo"
      Height          =   405
      Left            =   7230
      TabIndex        =   1
      Top             =   2610
      Width           =   1275
   End
   Begin VB.CommandButton cmNuevo 
      Caption         =   "&Nuevo Cálculo"
      Height          =   405
      Left            =   7230
      TabIndex        =   0
      Top             =   2070
      Width           =   1275
   End
   Begin MSDataGridLib.DataGrid xData 
      Height          =   4200
      Left            =   210
      TabIndex        =   8
      Top             =   165
      Width           =   6765
      _ExtentX        =   11933
      _ExtentY        =   7408
      _Version        =   393216
      AllowUpdate     =   0   'False
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
      Caption         =   "Planillas de Depósito de C.T.S."
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "Codigo"
         Caption         =   "Código"
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
         DataField       =   "Nombre"
         Caption         =   "Nombre Descriptivo"
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
         DataField       =   "Soles"
         Caption         =   "Total S/."
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
         DataField       =   "Dolares"
         Caption         =   "Total US$"
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
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         BeginProperty Column00 
            ColumnWidth     =   794.835
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   3195.213
         EndProperty
         BeginProperty Column02 
            Alignment       =   1
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnWidth     =   1005.165
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   390
      Left            =   195
      TabIndex        =   15
      Top             =   5160
      Width           =   2070
      _ExtentX        =   3651
      _ExtentY        =   688
      ButtonWidth     =   3228
      ButtonHeight    =   582
      TextAlignment   =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Otros Reportes        "
            Object.ToolTipText     =   "Click aquí para más reportes de trabajadores"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "CCOSTO"
                  Text            =   "Reporte de Deposito de CTS por Centro de Costo Detallado"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "TCTS"
                  Text            =   "Reporte Totales por Centro de Costo"
               EndProperty
            EndProperty
         EndProperty
      EndProperty
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "C.T.S."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   7890
      TabIndex        =   12
      Top             =   705
      Width           =   555
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "C.T.S."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   195
      Index           =   0
      Left            =   7920
      TabIndex        =   7
      Top             =   720
      Width           =   555
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   7980
      Picture         =   "frAdminCTS.frx":08CA
      Top             =   135
      Width           =   480
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000C&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   5640
      Left            =   60
      Top             =   90
      Width           =   7005
   End
End
Attribute VB_Name = "frAdminCTS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public xMes As String, xAnno As String
Dim WithEvents RSCTS As ADODB.Recordset
Attribute RSCTS.VB_VarHelpID = -1
Private Sub CMACEPTAR_CLICK()
    VPTRASPRM = "" & RSCTS!Codigo
    frAceptaCTS.Show 1
    RSCTS.Requery
    Set xData.DataSource = RSCTS
End Sub
Private Sub CMCALCMES_CLICK()
    Dim RSFORMULA As ADODB.Recordset
    Set RSFORMULA = New ADODB.Recordset
    RSFORMULA.Open "SELECT * FROM FORMULASCTS WHERE AFECTOPRO=1 ", DBSYSTEM, adOpenKeyset, adLockOptimistic
    If RSFORMULA.RecordCount > 0 Then
        If Len(Trim(RSFORMULA!CRITERIO)) > 0 Then
            xMes = Mid(RSFORMULA!CRITERIO, 1, 2)
            xAnno = Mid(RSFORMULA!CRITERIO, 3, 6)
        Else
            MsgBox "NO EXISTE CRITERIO DE COMPUTO EN LA FORMULA ACTIVA. PARA EL CALCULO MENSUAL ES NECESARIO EL CRITERIO DE COMPUTO " & Chr(13) & _
                    " SI UD. DESEA CALCULAR SIN CRITERIO IR A LA OPCION NUEVO DE CALCULO", vbInformation, "INFORMACION DEL SISTEMA"
            Exit Sub
        End If
    Else
        MsgBox "No existen Formulas para este Cálculo de CTS Mensual." & Chr(13) & "Verificar en menu BASE DE DATOS/OTROS ARCHIVOS/FORMULAS DE CTS"
        Exit Sub
    End If
    Set RSFORMULA = Nothing
    frCalcCTS2.Show 1
    RSCTS.Requery
End Sub
Private Sub cmCerrar_Click()
    Unload Me
End Sub
Private Sub CMCONSULTA_CLICK()
    VPTAREA = "VISTA"
    VPTRASPRM = "" & RSCTS!Codigo
    frCalcCTS.Show 1
End Sub
Private Sub CMCUSTODIA_CLICK()
    frCTSCustodia.Show 1
End Sub
Private Sub CMELIMINAR_CLICK()
    If RSCTS.EOF Or RSCTS.RecordCount = 0 Then
        MsgBox "No existe registros a eliminar", vbInformation
    Else
        MsgBox "ADVERTENCIA: Elimanara una Planilla de CTS, Sin posibilidad de Recuperar su Información", vbExclamation
        If MsgBox("Seguro de Eliminar la Planilla CTS : " & RSCTS!NOMBRE & " . SEGURO DE CONTINUAR", vbYesNo + vbQuestion) = vbNo Then Exit Sub
        DBSYSTEM.Execute "DELETE FROM CTS WHERE CODIGO=" & RSCTS!Codigo
        DBSYSTEM.Execute "DELETE FROM PLANCTS WHERE CODIGO=" & RSCTS!Codigo
        DBSYSTEM.Execute "DELETE FROM DETALLECTS WHERE CODIGO=" & RSCTS!Codigo
        RSCTS.Requery
        Set xData.DataSource = RSCTS
    End If
End Sub

Private Sub CMLISTADO_CLICK()
    Screen.MousePointer = 11
    If ExisteTablaAux(" [##TMPCTS1" & VGL_COMPUTER & "] ") Then DBSYSTEM.Execute "DROP TABLE  [##TMPCTS1" & VGL_COMPUTER & "] "
    DBSYSTEM.Execute "UPDATE PLANCTS SET BANCO='NONE' WHERE CTABANCO='' OR (CTABANCO)IS NULL"
    DBSYSTEM.Execute "SELECT PLANCTS.CODTRAB, LTRIM(APEPAT) + ' ' +  LTRIM(APEMAT) + ' ' + LTRIM(TRABAJADORES.NOMBRE) AS NOMBRES, PLANCTS.IMPORTECTS AS NETO, TIPDOC, DOCIDEN,PLANCTS.BANCO,PLANCTS.CTABANCO INTO  [##TMPCTS1" & VGL_COMPUTER & "]  FROM TRABAJADORES, PLANCTS WHERE PLANCTS.CODTRAB=TRABAJADORES.CODTRAB AND PLANCTS.CODIGO=" & RSCTS!Codigo & " AND CUSTODIA<>1"
    DBSYSTEM.Execute "UPDATE  [##TMPCTS1" & VGL_COMPUTER & "]  SET NETO=NETO"
    If ExisteTablaAux(" [##PAGOSXBANCO" & VGL_COMPUTER & "] ") Then DBSYSTEM.Execute "DROP TABLE  [##PAGOSXBANCO" & VGL_COMPUTER & "] "
    DBSYSTEM.Execute "SELECT * INTO  [##PAGOSXBANCO" & VGL_COMPUTER & "]  FROM  [##TMPCTS1" & VGL_COMPUTER & "]  "
    VPTRASPRM = "" & RSCTS!Codigo
    Screen.MousePointer = 1
    frPagoBcoCTS.Show 1
End Sub
Private Sub CMMODIFICAR_CLICK()
    VPTAREA = "MODIFICAR"
    VPTRASPRM = "" & RSCTS!Codigo
    frCalcCTS.Show 1
    RSCTS.Requery
    Set xData.DataSource = RSCTS
End Sub

Private Sub CMNUEVO_CLICK()
    VPTAREA = "NUEVO"
    frCalcCTS.Show 1
    RSCTS.Requery
    Set xData.DataSource = RSCTS
End Sub

Private Sub CMPRUEBA_CLICK()
    VPTAREA = "PRUEBA"
    frCalcCTS.Show 1
End Sub

Private Sub Command1_Click()
'    If RSCTS.RecordCount = 0 Then
'        MsgBox "No existen Registros a Imprimir", vbCritical
'        Exit Sub
'    End If
'    'REPORTE IMPRIMIR
'        With Reporte
'        .Reset
'        .WindowTitle = "PLAN0050 - PLANILLA DE PAGO DE C.T.S."
'        .ReportFileName = REGSISTEMA.REPORTES & "PLAN0050.RPT"
'        .Connect = "DSN=" & VGL_SERVERREP & ";User=sa;PWD=;DSQ=planilla_pp"
'        .StoredProcParam(0) = REGSISTEMA.BASESQL
'        .StoredProcParam(1) = "PLANCTS"
'        .Destination = crptToWindow
'        .WindowState = crptMaximized
'        .WindowShowPrintBtn = True
'        .WindowShowRefreshBtn = True
'        .WindowShowSearchBtn = True
'        .WindowShowPrintSetupBtn = True
'        .SortFields(0) = "+{SP_VISTA_DB.CODTRAB}"
'        .SelectionFormula = "{SP_VISTA_DB.CODIGO}=" & RSCTS!Codigo
'        .Formulas(0) = "XEMPRESA='" & REGSISTEMA.EMPRESA & "'"
'        .Formulas(1) = "XRUC='RUC N° " & REGSISTEMA.RUC & "'"
'        .Formulas(2) = "XMES='PLANILLA DE PAGOS DE C.T.S.: " & RSCTS!NOMBRE & "'"
'        If .Status <> 2 Then .Action = 1
'    End With
    Call ImprimirLista
End Sub
Private Sub ImprimirLista()
Dim SqlCad As String, TipCamb As String
    If ExisteTablaAux("##tmplistacts" & VGL_COMPUTER) Then
        DBAUXCOM.Execute "Drop Table " & "##tmplistacts" & VGL_COMPUTER
    End If
    TipCamb = MDIPrincipal.BarraEstado.Panels(3).Text
    
    SqlCad = "Select " & _
            "A.Codtrab,A.Nombres,C.DocIden,C.FechaNac,C.CTABANCO,C.CTACTS, " & _
            "TOTALREM=B.REMUNAFEC,A.IMPORTECTS,MON=case when Isnull(rtrim(ltrim(c.MON)),'')='' then '01' else c.MON end, " & _
            "TIPCTA = " & TipCamb & "   " & _
            "INTO ##tmplistacts" & VGL_COMPUTER & "  " & _
            "From PLANCTS A, " & _
            "(Select CODTRAB,REMUNAFEC=SUM(IMPORTE) From dbo.DETALLECTS " & _
            "Where Codigo = " & RSCTS!Codigo & "  " & _
            "GROUP BY CODTRAB) AS B, " & _
            "TRABAJADORES C " & _
            "Where A.Codtrab=b.Codtrab and " & _
            "      A.Codtrab=C.Codtrab and " & _
            "      A.Codigo = " & RSCTS!Codigo
            
     DBSYSTEM.Execute SqlCad
    With Reporte
        .Reset
        .WindowTitle = "pl_listactses - PLANILLA DE PAGO DE C.T.S."
        .ReportFileName = REGSISTEMA.REPORTES & "pl_listactses.rpt"
        .Connect = "DSN=" & VGL_SERVERREP & ";User=sa;PWD=;DSQ=planilla_pp"
        .StoredProcParam(0) = "##tmplistacts" & VGL_COMPUTER
        
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .WindowShowPrintBtn = True
        .WindowShowRefreshBtn = True
        .WindowShowSearchBtn = True
        .WindowShowPrintSetupBtn = True
        '.Formulas(0) = "XEMPRESA='" & REGSISTEMA.EMPRESA & "'"
        '.Formulas(1) = "XRUC='RUC N° " & REGSISTEMA.RUC & "'"
        .Formulas(2) = "XMES='" & RSCTS!NOMBRE & "'"
        .DiscardSavedData = True
        If .Status <> 2 Then .Action = 1
    End With
End Sub

Private Sub Command2_Click()
    VPTAREA = "REPORTE"
    If RSCTS.RecordCount > 0 Then
        VPTRASPRM = "" & RSCTS!Codigo
        FrPerCts.Show 1
    Else: MsgBox "No existen registros en la Tabla Liquidación de CTS", vbCritical
    End If
End Sub
Private Sub Form_Load()
    Dim RSFOR As New ADODB.Recordset
    RSFOR.Open "SELECT * FROM FORMULASCTS WHERE AFECTOPRO=1", DBSYSTEM, adOpenKeyset, adLockOptimistic
    If RSFOR.RecordCount > 0 Then
       If Len(Trim(RSFOR!CRITERIO)) > 0 Then
            cmCalcMes.Enabled = True
            cmNuevo.Enabled = False
            cmPrueba.Enabled = False
            cmModificar.Enabled = False
        Else
            cmCalcMes.Enabled = False
            cmNuevo.Enabled = True
            cmPrueba.Enabled = True
            cmModificar.Enabled = True
        End If
    End If
    Set RSFOR = Nothing
    Set RSCTS = New ADODB.Recordset
    DBSYSTEM.Execute "UPDATE PLANCTS SET CUSTODIA=0 WHERE (CUSTODIA)IS NULL"
    RSCTS.Open "CTS", DBSYSTEM, adOpenKeyset, adLockOptimistic
    Set xData.DataSource = RSCTS
End Sub

Private Sub FORM_UNLOAD(CANCEL As Integer)
    Set RSCTS = Nothing
End Sub

Private Sub RSCTS_MOVECOMPLETE(ByVal ADREASON As ADODB.EventReasonEnum, ByVal PERROR As ADODB.Error, ADSTATUS As ADODB.EventStatusEnum, ByVal PRECORDSET As ADODB.Recordset)
    If RSCTS.EOF Or RSCTS.RecordCount = 0 Or RSCTS.BOF Then
        cmAceptar.Enabled = False
        cmCustodia.Enabled = False
        cmEliminar.Enabled = False
        cmListado.Enabled = False
        cmModificar.Enabled = False
        cmConsulta.Enabled = False
    Else
        If RSCTS!CERRADO = 1 Then
            cmAceptar.Enabled = False
        Else
            cmAceptar.Enabled = True
        End If
        cmCustodia.Enabled = True
        cmEliminar.Enabled = True
        cmListado.Enabled = True
        cmModificar.Enabled = True
        cmConsulta.Enabled = True
    End If
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Dim X As Integer, Y As Integer
Dim SQL As String
Dim RSAUX As New ADODB.Recordset
            If ExisteTablaAux(" [##_TMPCTSREP" & VGL_COMPUTER & "] ") Then DBSYSTEM.Execute "DROP TABLE  [##_TMPCTSREP" & VGL_COMPUTER & "] "
                SQL = "SELECT PLANCTS.CODIGO, PLANCTS.CODTRAB, PLANCTS.NOMBRES, PLANCTS.IMPORTECTS, VWTRABAJ.CODCCOSTO, VWTRABAJ.CENTRO, VWTRABAJ.FECHAING, VWTRABAJ.BASICO" & _
                    " INTO  [##_TMPCTSREP" & VGL_COMPUTER & "]  FROM VWTRABAJ, PLANCTS WHERE VWTRABAJ.CODTRAB = PLANCTS.CODTRAB AND CODIGO=" & RSCTS!Codigo
                    DBSYSTEM.Execute SQL, X
                            For Y = 0 To X
                            Next
    Select Case ButtonMenu.KEY
        Case "CCOSTO"
            CambiaPanelBD True
            Screen.MousePointer = vbHourglass
                With Reporte
                    .Reset
                    .WindowTitle = "PLAN0089 - REPORTE DE CALCULO MENSUAL DE CTS POR CENTRO DE COSTO DETALLADO"
                    .ReportFileName = REGSISTEMA.REPORTES & "\PLAN0089.RPT"
                    .Connect = "DSN=" & VGL_SERVERREP & ";User=sa;PWD=;DSQ=planilla_pp"
                    .StoredProcParam(0) = " [##_TMPCTSREP" & VGL_COMPUTER & "] "
                    .Destination = crptToWindow
                    .WindowState = crptMaximized
                    .WindowShowPrintBtn = True
                    .WindowShowRefreshBtn = True
                    .WindowShowSearchBtn = True
                    .WindowShowPrintSetupBtn = True
                    .Formulas(0) = "XEMPRESA='" & REGSISTEMA.EMPRESA & "'"
                    RSAUX.Open "SELECT * FROM CTS WHERE CODIGO=" & RSCTS!Codigo, DBSYSTEM, adOpenKeyset, adLockOptimistic
                        If RSAUX.RecordCount > 0 Then
                            .Formulas(1) = "XPERIODO='" & RSAUX!NOMBRE & "'"
                        Else
                            .Formulas(1) = "XPERIODO='" & Format(Date, "MMMM - YYYY") & "'"
                        End If
                    Set RSAUX = Nothing
                    If .Status <> 2 Then .Action = 1
                End With
                CambiaPanelBD False
                Screen.MousePointer = vbNormal
        Case "TCTS"
            CambiaPanelBD True
            Screen.MousePointer = vbHourglass
                'VERIFICAR KOKI
                With Reporte
                    .Reset
                    .WindowTitle = "PLAN0090 - REPORTE DE CALCULO MENSUAL DE CTS POR CENTRO DE COSTO"
                    .ReportFileName = REGSISTEMA.REPORTES & "\PLAN0090.RPT"
                    .Connect = "DSN=" & VGL_SERVERREP & ";User=sa;PWD=;DSQ=planilla_pp"
                    .StoredProcParam(0) = " [##_TMPCTSREP" & VGL_COMPUTER & "] "
                    .Destination = crptToWindow
                    .WindowState = crptMaximized
                    .WindowShowPrintBtn = True
                    .WindowShowRefreshBtn = True
                    .WindowShowSearchBtn = True
                    .WindowShowPrintSetupBtn = True
                    .Formulas(0) = "XEMPRESA='" & REGSISTEMA.EMPRESA & "'"
                    RSAUX.Open "SELECT * FROM CTS WHERE CODIGO=" & RSCTS!Codigo, DBSYSTEM, adOpenKeyset, adLockOptimistic
                        If RSAUX.RecordCount > 0 Then
                            .Formulas(1) = "XPERIODO='" & RSAUX!NOMBRE & "'"
                        Else
                            .Formulas(1) = "XPERIODO='" & Format(Date, "MMMM - YYYY") & "'"
                        End If
                    Set RSAUX = Nothing
                    If .Status <> 2 Then .Action = 1
                End With
                CambiaPanelBD False
                Screen.MousePointer = vbNormal
        End Select
End Sub

