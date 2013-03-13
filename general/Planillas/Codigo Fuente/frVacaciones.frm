VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frVacaciones 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vacaciones de los Trabajadores"
   ClientHeight    =   5100
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7785
   ForeColor       =   &H00000080&
   Icon            =   "frVacaciones.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   7785
   Tag             =   "Administración de Vacaciones "
   Begin Crystal.CrystalReport Reporte 
      Left            =   6420
      Top             =   1335
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00808080&
      Caption         =   "Todos los meses del año"
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   3885
      TabIndex        =   18
      Top             =   270
      Width           =   2070
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   390
      Left            =   6315
      TabIndex        =   17
      Top             =   4035
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   688
      ButtonWidth     =   2011
      ButtonHeight    =   582
      TextAlignment   =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "  Otros      "
            Key             =   "Reportes"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   4
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "RolGeneral"
                  Text            =   "Rol de Vacaciones General"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "RolMes"
                  Text            =   "Rol de Vacaciones por Mes"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Salida"
                  Text            =   "Salida de Vacaciones"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Penvac"
                  Text            =   "Pendientes Vacaciones"
               EndProperty
            EndProperty
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
   Begin VB.OptionButton Option4 
      BackColor       =   &H00808080&
      Caption         =   "Por Goce"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   4965
      TabIndex        =   16
      Top             =   690
      Width           =   1050
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H00808080&
      Caption         =   "Por Traspaso"
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   3360
      TabIndex        =   15
      Top             =   690
      Width           =   1335
   End
   Begin VB.CommandButton cmProgramacion 
      Caption         =   "Programar Vacaciones"
      Height          =   465
      Left            =   4695
      TabIndex        =   14
      Top             =   4455
      Width           =   1320
   End
   Begin VB.CommandButton cmAutomatico 
      Caption         =   "Agregar Automático"
      Height          =   465
      Left            =   1695
      TabIndex        =   13
      Top             =   4455
      Width           =   1320
   End
   Begin VB.CommandButton cmGoce 
      Caption         =   "Goce de Vacaciones"
      Height          =   465
      Left            =   3195
      TabIndex        =   12
      Top             =   4455
      Width           =   1320
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00808080&
      Caption         =   "Canceladas"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   1875
      TabIndex        =   9
      Top             =   660
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00808080&
      Caption         =   "Programadas"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   270
      TabIndex        =   8
      Top             =   675
      Value           =   -1  'True
      Width           =   1380
   End
   Begin MSComCtl2.DTPicker xMes 
      Height          =   300
      Left            =   900
      TabIndex        =   7
      Top             =   210
      Width           =   2205
      _ExtentX        =   3889
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "MMMM' de' yyyy"
      Format          =   61931523
      CurrentDate     =   36844
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Reporte"
      Height          =   330
      Left            =   6315
      TabIndex        =   5
      Top             =   3555
      Width           =   1335
   End
   Begin VB.CommandButton cmCerrar 
      Cancel          =   -1  'True
      Caption         =   "&Cerrar"
      Height          =   330
      Left            =   6315
      TabIndex        =   4
      Top             =   4590
      Width           =   1335
   End
   Begin VB.CommandButton cmEliminar 
      Caption         =   "&Eliminar"
      Height          =   330
      Left            =   6315
      TabIndex        =   3
      Top             =   3060
      Width           =   1335
   End
   Begin VB.CommandButton cmModificar 
      Caption         =   "&Modificar"
      Height          =   330
      Left            =   6315
      TabIndex        =   2
      Top             =   2565
      Width           =   1335
   End
   Begin VB.CommandButton cmAgregar 
      Caption         =   "&Agregar Personalizado"
      Height          =   465
      Left            =   195
      TabIndex        =   1
      Top             =   4455
      Width           =   1320
   End
   Begin MSDataGridLib.DataGrid xData 
      Height          =   3420
      Left            =   195
      TabIndex        =   0
      Top             =   975
      Width           =   5820
      _ExtentX        =   10266
      _ExtentY        =   6033
      _Version        =   393216
      BackColor       =   14869218
      HeadLines       =   2
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
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vacaciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   1
      Left            =   6570
      TabIndex        =   11
      Top             =   780
      Width           =   1005
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Vacaciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   6570
      TabIndex        =   10
      Top             =   765
      Width           =   1005
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   7080
      Picture         =   "frVacaciones.frx":08CA
      Top             =   255
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   6810
      Picture         =   "frVacaciones.frx":1194
      Top             =   165
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mes"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   315
      TabIndex        =   6
      Top             =   270
      Width           =   300
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      Height          =   4875
      Left            =   90
      Top             =   135
      Width           =   6045
   End
End
Attribute VB_Name = "frVacaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RSTRABS As New ADODB.Recordset
Dim REGACT As REGWIN
Dim SQLREPORT As String
Private Sub CHECK1_CLICK()
    XMES_CHANGE
End Sub

Private Sub CMAGREGAR_CLICK()
    VPTAREA = "NUEVO"
    frCancelVacaciones.Show 1
    XMES_CHANGE
End Sub

Private Sub CMAUTOMATICO_CLICK()
    If Not Option1.Value Then Exit Sub
    If RSTRABS.EOF Then
        MsgBox "Deberá de existir o seleccionar un trabajador con Programación de Vacaciones. El Cálculo automático de Vacaciones se realiza a través de una programación previa del periodo vacacional.", vbInformation
        Exit Sub
    End If
    Load frCalcVac2
    With frCalcVac2
        .Frame1.Tag = RSTRABS!Codigo
        .xTrab.Tag = "" & RSTRABS!CODTRAB
        .xTrab.Caption = "" & RSTRABS!NOMBRES
        .xSalini.Value = RSTRABS!FECHAINI
        .xSalFin.Value = RSTRABS!FECHAFIN
        .xFecha.Caption = DevuelveValor("SELECT FECHAING FROM TRABAJADORES WHERE CODTRAB='" & RSTRABS!CODTRAB & "'", DBSYSTEM)
        .xFechaIni = DevuelveValor("SELECT FECHAINICAL FROM HISTOVAC WHERE CODIGO=" & RSTRABS!Codigo, DBSYSTEM)
        .xFechaFin = DevuelveValor("SELECT FECHAFINCAL FROM HISTOVAC WHERE CODIGO=" & RSTRABS!Codigo, DBSYSTEM)
        .xMeses.Text = 12
        .xDias.Text = 0
        .Show 1
    End With
    XMES_CHANGE
End Sub

Private Sub cmCerrar_Click()
    Unload Me
End Sub

Private Sub CMELIMINAR_CLICK()
    If RSTRABS.EOF Then Exit Sub
    Select Case xMes.Tag
        Case 1
            If MsgBox("Confirma que desea eliminar el registro de vacaciones seleccionado", vbYesNo + vbQuestion) = vbNo Then Exit Sub
            DBSYSTEM.Execute "DELETE FROM HISTOVAC WHERE CODIGO=" & RSTRABS!Codigo
            DBSYSTEM.Execute "DELETE FROM DETALLEVAC WHERE CODIGO=" & RSTRABS!Codigo
        Case 3
            'TIPO TRASPASO
            If MsgBox("Realmente desea eliminar el calculo de vacaciones seleccionado", vbYesNo + vbQuestion) = vbNo Then Exit Sub
            DBSYSTEM.Execute "UPDATE HISTOVAC SET MONTO=0, NOMBOL=0 WHERE CODIGO=" & RSTRABS!Codigo
            DBSYSTEM.Execute "DELETE FROM DETALLEVAC WHERE CODIGO=" & RSTRABS!Codigo
        Case 4
    End Select
    XMES_CHANGE
End Sub

Private Sub CMGOCE_CLICK()
    Load frmDiaEspecial
    With frmDiaEspecial
        If RSTRABS.RecordCount > 0 Then
            .xTrab.Tag = RSTRABS!CODTRAB
            .xTrab.Text = RSTRABS!NOMBRES
            .xPendiente.Caption = RSTRABS!Dias
            .xPendiente.Tag = RSTRABS!Codigo
        End If
        .Show 1
    End With
    XMES_CHANGE
End Sub

Private Sub CMMODIFICAR_CLICK()
    If RSTRABS.EOF Or RSTRABS.RecordCount = 0 Then Exit Sub
    Select Case xMes.Tag
        Case 1
            'PROGRAMACIÓN DE VACACIONES
            Load frPrgVac2
            With frPrgVac2
                VPTAREA = RSTRABS!Codigo
                .xTrab.Tag = RSTRABS!CODTRAB
                .xTrab.Text = "" & RSTRABS!NOMBRES
                .xSalini.Value = RSTRABS!FECHAINI
                .xSalFin.Value = RSTRABS!FECHAFIN
                .xPerIni.Value = DevuelveValor("SELECT FECHAINICAL FROM HISTOVAC WHERE CODIGO=" & RSTRABS!Codigo, DBSYSTEM)
                .xPerFin.Value = DevuelveValor("SELECT FECHAFINCAL FROM HISTOVAC WHERE CODIGO=" & RSTRABS!Codigo, DBSYSTEM)
                .xArea.Caption = DevuelveValor("SELECT NOMBREAREA FROM VWTRABAJ WHERE CODTRAB='" & RSTRABS!CODTRAB & "'", DBSYSTEM)
                .Show 1
            End With
        Case 2
        Case 3
            If DevuelveValor("SELECT MODOCALCULO FROM HISTOVAC WHERE CODIGO=" & RSTRABS!Codigo, DBSYSTEM) = 1 Then
                'SI HA SIDO EN FORMA MANUAL
                VPTAREA = "" & RSTRABS!Codigo
                frCancelVacaciones.Show 1
            Else
                Load frCalcVac2
                With frCalcVac2
                    .xTrab.Caption = RSTRABS!NOMBRES
                    .xTrab.Tag = RSTRABS!CODTRAB
                    .xFecha.Caption = DevuelveValor("SELECT FECHAING FROM TRABAJADORES WHERE CODTRAB='" & RSTRABS!CODTRAB & "'", DBSYSTEM)
                    .Show 1
                End With
            End If
        Case 4
    End Select
    XMES_CHANGE
End Sub

Private Sub CMPROGRAMACION_CLICK()
    If Option1.Value Then
        VPTAREA = "NUEVO"
        frPrgVac2.Show 1
        XMES_CHANGE
    End If
End Sub

Private Sub Command1_Click()
    'REPORTE IMPRIMIR
    Dim FEC1 As Date
    xMes.Day = 1
    FEC1 = DateAdd("D", -1, DateAdd("M", 1, xMes.Value))
    DBSTARPLAN.Execute "EXECUTE SP_VACACIONES '" & REGSISTEMA.BASESQL & "', '" & Year(xMes.Value) & "', " & DateSQL(xMes.Value) & ", " & DateSQL(FEC1) & ", " & xMes.Tag & ", " & Check1.Value & ", '" & VGL_COMPUTER & "'"
    Dim ReportFileName  As String
    Select Case xMes.Tag
        Case 1
            ReportFileName = REGSISTEMA.REPORTES & "PLAN0069.RPT"
        Case 2
         '  ReportFileName = REGSISTEMA.REPORTES & "PLAN0070.RPT"
        Case 3
            ReportFileName = REGSISTEMA.REPORTES & "PLAN0071.RPT"
        Case 4
            ReportFileName = REGSISTEMA.REPORTES & "PLAN0075.RPT"
    End Select
        With Reporte
            .Reset
            .ReportFileName = ReportFileName
             '.LogOnServer "pdssql.dll", VGL_SERVERREP, "MARFICE_PP", "SOPORTE", "SOPORTE"
            .Connect = "DSN=" & VGL_SERVERREP & ";User=sa;PWD=;DSQ=" & VGL_BASE & ""
            .StoredProcParam(0) = "##TMPVAC001" & VGL_COMPUTER & ""
            .Destination = crptToWindow
            .WindowShowPrintBtn = True
            .WindowShowRefreshBtn = True
            .WindowShowSearchBtn = True
            .WindowShowPrintSetupBtn = True
            .WindowState = crptMaximized
            .WindowTitle = .ReportFileName
            .Formulas(0) = "XEMPRESA='" & REGSISTEMA.EMPRESA & "'"
            .Formulas(1) = "XPERIODO='" & IIf(Check1.Value = 1, "AÑO " & Year(xMes.Value), "MES DE " & AMESES(xMes.Month) & " DEL " & Year(xMes.Value)) & "'"
             If .Status <> 2 Then .Action = 1
    End With
End Sub
Private Sub IMPRIMIR()
    Dim arrform(2) As Variant, arrparm(2) As Variant
    '@BASE, @CODTRAB, @GRUPO, @TIPO, @TipoTrab, @flagfecha, @Fechaini, @fechafin
    arrparm(0) = REGSISTEMA.BASESQL
    arrparm(1) = Month(xMes)
    arrform(0) = "XPERIODO='" & IIf(Check1.Value = 1, "AÑO " & Year(xMes.Value), "MES DE " & AMESES(xMes.Month) & " DEL " & Year(xMes.Value)) & "'"
    Call ImpresionRptProc("pl_vagozo.rpt", arrform, arrparm, , "Listado Pendientes de Vacaciones")
End Sub

Private Sub Form_Activate()
    ActivarTools REGACT
    xMes.Value = Date
End Sub

Private Sub Form_Load()
    With REGACT
        .BUSCAR = False
        .EDITAR = True
        .ELIMINAR = True
        .FILTRAR = False
        .IMPRIMIR = True
        .NUEVO = True
        .PRELIMINAR = True
    End With
    OPTION1_CLICK
End Sub

Private Sub FORM_UNLOAD(CANCEL As Integer)
    Set RSTRABS = Nothing
End Sub

Private Sub OPTION1_CLICK()
    'PROGRAMACIÓN DE VACACIONES
    DBSYSTEM.Execute "UPDATE HISTOVAC SET MONTO=0 WHERE (MONTO) IS NULL"
    Dim FEC1 As Date
    xMes.Day = 1
    FEC1 = DateAdd("D", -1, DateAdd("M", 1, xMes.Value))
    Set RSTRABS = Nothing
    If Check1.Value = 1 Then
        RSTRABS.Open "SELECT CODIGO,HISTOVAC.CODTRAB,NOMBRES,FECHAINI, FECHAFIN FROM HISTOVAC,VWTRABAJ WHERE HISTOVAC.CODTRAB=VWTRABAJ.CODTRAB AND CERRADO=0 AND PROGRAMADO=1 AND MONTO=0 AND YEAR(FECHAINI)=" & Year(xMes.Value), DBSYSTEM, adOpenStatic, adLockReadOnly
    Else
        RSTRABS.Open "SELECT CODIGO,HISTOVAC.CODTRAB,NOMBRES,FECHAINI, FECHAFIN FROM HISTOVAC,VWTRABAJ WHERE HISTOVAC.CODTRAB=VWTRABAJ.CODTRAB AND CERRADO=0 AND PROGRAMADO=1 AND MONTO=0 AND (FECHAINI BETWEEN " & DateSQL(xMes.Value) & " AND " & DateSQL(FEC1) & ")", DBSYSTEM, adOpenStatic, adLockReadOnly
    End If
    Set xData.DataSource = RSTRABS
    With xData
        .Columns("NOMBRES").Width = 3000
        .Columns("FECHAINI").Width = 1100
        .Columns("FECHAINI").Caption = "FECHA DE INICIO"
        .Columns("FECHAFIN").Width = 1100
        .Columns("FECHAFIN").Caption = "FECHA FINAL"
        .Columns("CODIGO").Visible = False
        .Columns("CODTRAB").Visible = False
    End With
    TODASDISABLED
    cmAutomatico.Enabled = True
    cmProgramacion.Enabled = True
    xMes.Tag = 1
End Sub

Private Sub OPTION2_Click()
    'CANCELADAS
    Dim FEC1 As Date
    xMes.Day = 1
    FEC1 = DateAdd("D", -1, DateAdd("M", 1, xMes.Value))
    Set RSTRABS = Nothing
    If Check1.Value = 1 Then
        RSTRABS.Open "SELECT CODIGO,HISTOVAC.CODTRAB,NOMBRES,MONTO, FECHAINI, FECHAFIN FROM HISTOVAC,VWTRABAJ WHERE HISTOVAC.CODTRAB=VWTRABAJ.CODTRAB AND NOMBOL<>0 AND CERRADO=1 AND YEAR(FECHAINI)=" & Year(xMes.Value), DBSYSTEM, adOpenStatic, adLockReadOnly
    Else
        RSTRABS.Open "SELECT CODIGO,HISTOVAC.CODTRAB,NOMBRES,MONTO, FECHAINI, FECHAFIN FROM HISTOVAC,VWTRABAJ WHERE HISTOVAC.CODTRAB=VWTRABAJ.CODTRAB AND NOMBOL<>0 AND CERRADO=1 AND (FECHAINI BETWEEN " & DateSQL(xMes.Value) & " AND " & DateSQL(FEC1) & ")", DBSYSTEM, adOpenStatic, adLockReadOnly
    End If
    Set xData.DataSource = RSTRABS
    With xData
        .Columns("NOMBRES").Width = 3000
        .Columns("FECHAINI").Width = 1100
        .Columns("FECHAINI").Caption = "FECHA DE INICIO"
        .Columns("FECHAFIN").Width = 1100
        .Columns("FECHAFIN").Caption = "FECHA FINAL"
        .Columns("CODIGO").Visible = False
        .Columns("MONTO").Alignment = dbgRight
        .Columns("MONTO").NumberFormat = "0.00 "
        .Columns("CODTRAB").Visible = False
    End With
    TODASDISABLED
    xMes.Tag = 2
End Sub

Private Sub OPTION3_Click()
    'PENDIENTE
    Dim FEC1 As Date
    xMes.Day = 1
    FEC1 = DateAdd("D", -1, DateAdd("M", 1, xMes.Value))
    Set RSTRABS = Nothing
    If Check1.Value = 1 Then
        RSTRABS.Open "SELECT CODIGO,HISTOVAC.CODTRAB,NOMBRES,MONTO, FECHAINI, FECHAFIN FROM HISTOVAC,VWTRABAJ WHERE HISTOVAC.CODTRAB=VWTRABAJ.CODTRAB AND CERRADO=0 AND MONTO<>0 AND YEAR(FECHAINI)=" & Year(xMes.Value), DBSYSTEM, adOpenStatic, adLockReadOnly
    Else
        RSTRABS.Open "SELECT CODIGO,HISTOVAC.CODTRAB,NOMBRES,MONTO, FECHAINI, FECHAFIN FROM HISTOVAC,VWTRABAJ WHERE HISTOVAC.CODTRAB=VWTRABAJ.CODTRAB AND CERRADO=0 AND MONTO<>0 AND (FECHAINI BETWEEN " & DateSQL(xMes.Value) & " AND " & DateSQL(FEC1) & ")", DBSYSTEM, adOpenStatic, adLockReadOnly
    End If
    Set xData.DataSource = RSTRABS
    With xData
        .Columns("NOMBRES").Width = 3000
        .Columns("FECHAINI").Width = 1100
        .Columns("FECHAINI").Caption = "FECHA DE INICIO"
        .Columns("FECHAFIN").Width = 1100
        .Columns("FECHAFIN").Caption = "FECHA FINAL"
        .Columns("CODIGO").Visible = False
        .Columns("MONTO").Alignment = dbgRight
        .Columns("MONTO").NumberFormat = "0.00 "
        .Columns("CODTRAB").Visible = False
    End With
    TODASDISABLED
    cmAgregar.Enabled = True
    xMes.Tag = 3
End Sub

Private Sub OPTION4_Click()
    'GOCE VACACIONES
    Dim FEC1 As Date
    xMes.Day = 1
    FEC1 = DateAdd("D", -1, DateAdd("M", 1, xMes.Value))
    Set RSTRABS = Nothing
    RSTRABS.Open "SELECT CODIGO,HISTOVAC.CODTRAB,NOMBRES,DIAS, FECHAINI, FECHAFIN FROM HISTOVAC,VWTRABAJ WHERE HISTOVAC.CODTRAB=VWTRABAJ.CODTRAB AND DIAS<>0 AND NOMBOL<>0 AND CERRADO=1", DBSYSTEM, adOpenStatic, adLockReadOnly
    Set xData.DataSource = RSTRABS
    With xData
        .Columns("NOMBRES").Width = 3000
        .Columns("FECHAINI").Width = 1100
        .Columns("FECHAINI").Caption = "FECHA DE INICIO"
        .Columns("FECHAFIN").Width = 1100
        .Columns("FECHAFIN").Caption = "FECHA FINAL"
        .Columns("CODIGO").Visible = False
        .Columns("DIAS").Alignment = dbgCenter
        .Columns("DIAS").Width = 600
        .Columns("CODTRAB").Visible = False
    End With
    TODASDISABLED
    cmGoce.Enabled = True
    xMes.Tag = 4
End Sub

Public Sub TODASDISABLED()
    cmAgregar.Enabled = False
    cmAutomatico.Enabled = False
    cmGoce.Enabled = False
    cmProgramacion.Enabled = False
End Sub


Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Dim FEC1

If ButtonMenu.INDEX = 4 Then
    Call IMPRIMIR
    Exit Sub
End If



FEC1 = DateAdd("d", -1, DateAdd("m", 1, xMes.Value))
If ExisteTablaAux(" [##TMPVAC" & VGL_COMPUTER & "] ") = True Then DBSYSTEM.Execute "DROP TABLE  [##TMPVAC" & VGL_COMPUTER & "] "
If Option1.Value = True Then
    If ButtonMenu.INDEX = 1 Then
        DBSYSTEM.Execute "SELECT CODIGO, HISTOVAC.CODTRAB, NOMBRES, FECHAINI, FECHAFIN INTO  [##TMPVAC" & VGL_COMPUTER & "]  FROM " & REGSISTEMA.BASESQL & ".dbo.HISTOVAC HISTOVAC, " & REGSISTEMA.BASESQL & ".dbo.VWTRABAJ VWTRABAJ WHERE HISTOVAC.CODTRAB = VWTRABAJ.CODTRAB AND CERRADO=0 AND PROGRAMADO=1 AND MONTO=0 AND YEAR(FECHAINI)=" & Year(xMes.Value)
    Else
        DBSYSTEM.Execute "SELECT CODIGO, HISTOVAC.CODTRAB, NOMBRES, FECHAINI, FECHAFIN INTO  [##TMPVAC" & VGL_COMPUTER & "]  FROM " & REGSISTEMA.BASESQL & ".dbo.HISTOVAC HISTOVAC, " & REGSISTEMA.BASESQL & ".dbo.VWTRABAJ VWTRABAJ WHERE HISTOVAC.CODTRAB = VWTRABAJ.CODTRAB AND CERRADO=0 AND PROGRAMADO=1 AND MONTO=0 AND (FECHAINI BETWEEN " & DateSQL(xMes.Value) & " AND " & DateSQL(FEC1) & ")"
    End If
    
    With Reporte
        .Reset
        .Connect = "DSN=" & VGL_SERVERREP & ";User=sa;PWD=;DSQ=planilla_pp"
        .ReportFileName = REGSISTEMA.REPORTES & "PLAN0082.RPT"
        .StoredProcParam(0) = " [##TMPVAC" & VGL_COMPUTER & "] "
        .Destination = crptToWindow
        .WindowShowPrintBtn = True
        .WindowShowRefreshBtn = True
        .WindowShowSearchBtn = True
        .WindowShowPrintSetupBtn = True
        .WindowState = crptMaximized
        .WindowTitle = .ReportFileName
        .Formulas(0) = "Empresa='" & REGSISTEMA.EMPRESA & "'"
        .Formulas(1) = "Titulo='" & UCase(ButtonMenu.Text) & " " & UCase(Option1.Caption) & "'"
        If .Status <> 2 Then .Action = 1
    End With
    
End If

If Option2.Value = True Then
    If ButtonMenu.INDEX = 0 Then
        DBSYSTEM.Execute "SELECT CODIGO, HISTOVAC.CODTRAB, NOMBRES, MONTO, FECHAINI, FECHAFIN INTO  [##TMPVAC" & VGL_COMPUTER & "]  FROM " & REGSISTEMA.BASESQL & ".dbo.HISTOVAC HISTOVAC, " & REGSISTEMA.BASESQL & ".dbo.VWTRABAJ VWTRABAJ WHERE HISTOVAC.CODTRAB = VWTRABAJ.CODTRAB AND NOMBOL<>0 AND CERRADO=1 AND YEAR(FechaIni)=" & Year(xMes.Value)
    Else
        DBSYSTEM.Execute "SELECT CODIGO, HISTOVAC.CODTRAB, NOMBRES, MONTO, FECHAINI, FECHAFIN INTO  [##TMPVAC" & VGL_COMPUTER & "]  FROM " & REGSISTEMA.BASESQL & ".dbo.HISTOVAC HISTOVAC, " & REGSISTEMA.BASESQL & ".dbo.VWTRABAJ VWTRABAJ WHERE HISTOVAC.CODTRAB = VWTRABAJ.CODTRAB AND NOMBOL<>0 AND CERRADO=1 AND (FECHAINI BETWEEN " & DateSQL(xMes.Value) & " AND " & DateSQL(FEC1) & ")"
    End If
    
    With Reporte
        .Reset
        .Connect = "DSN=" & VGL_SERVERREP & ";User=sa;PWD=;DSQ=planilla_pp"
        .ReportFileName = REGSISTEMA.REPORTES & "PLAN0083.RPT"
        .StoredProcParam(0) = " [##TMPVAC" & VGL_COMPUTER & "] "
        .Destination = crptToWindow
        .WindowShowPrintBtn = True
        .WindowShowRefreshBtn = True
        .WindowShowSearchBtn = True
        .WindowShowPrintSetupBtn = True
        .WindowState = crptMaximized
        .WindowTitle = .ReportFileName
        .Formulas(0) = "Empresa='" & REGSISTEMA.EMPRESA & "'"
        .Formulas(1) = "Titulo='" & UCase(ButtonMenu.Text) & " " & UCase(Option2.Caption) & "'"
        If .Status <> 2 Then .Action = 1
    End With
    
End If

If Option3.Value = True Then
    If ButtonMenu.INDEX = 0 Then
        DBSYSTEM.Execute "SELECT CODIGO,HISTOVAC.CODTRAB, NOMBRES, MONTO, FECHAINI, FECHAFIN INTO  [##TMPVAC" & VGL_COMPUTER & "]  FROM " & REGSISTEMA.BASESQL & ".dbo.HISTOVAC HISTOVAC, " & REGSISTEMA.BASESQL & ".dbo.VWTRABAJ VWTRABAJ WHERE HISTOVAC.CODTRAB = VWTRABAJ.CODTRAB AND CERRADO=0 AND MONTO<>0 AND YEAR(FECHAINI)=" & Year(xMes.Value)
    Else
        DBSYSTEM.Execute "SELECT CODIGO,HISTOVAC.CODTRAB, NOMBRES, MONTO, FECHAINI, FECHAFIN INTO  [##TMPVAC" & VGL_COMPUTER & "]  FROM " & REGSISTEMA.BASESQL & ".dbo.HISTOVAC HISTOVAC, " & REGSISTEMA.BASESQL & ".dbo.VWTRABAJ VWTRABAJ WHERE HISTOVAC.CODTRAB = VWTRABAJ.CODTRAB AND CERRADO=0 AND MONTO<>0 AND (FECHAINI BETWEEN " & DateSQL(xMes.Value) & " AND " & DateSQL(FEC1) & ")"
    End If
    
    With Reporte
        .Reset
        .Connect = "DSN=" & VGL_SERVERREP & ";User=sa;PWD=;DSQ=planilla_pp"
        .ReportFileName = REGSISTEMA.REPORTES & "PLAN0084.RPT"
        .StoredProcParam(0) = " [##TMPVAC" & VGL_COMPUTER & "] "
        .Destination = crptToWindow
        .WindowShowPrintBtn = True
        .WindowShowRefreshBtn = True
        .WindowShowSearchBtn = True
        .WindowShowPrintSetupBtn = True
        .WindowState = crptMaximized
        .WindowTitle = .ReportFileName
        .Formulas(0) = "Empresa='" & REGSISTEMA.EMPRESA & "'"
        .Formulas(1) = "Titulo='" & UCase(ButtonMenu.Text) & " " & UCase(Option3.Caption) & "'"
        If .Status <> 2 Then .Action = 1
    End With
End If
If Option4.Value = True Then
    DBSYSTEM.Execute "SELECT CODIGO, HISTOVAC.CODTRAB, NOMBRES, DIAS, FECHAINI, FECHAFIN INTO  [##TMPVAC" & VGL_COMPUTER & "]  FROM " & REGSISTEMA.BASESQL & ".dbo.HISTOVAC HISTOVAC, " & REGSISTEMA.BASESQL & ".dbo.VWTRABAJ VWTRABAJ WHERE HISTOVAC.CODTRAB = VWTRABAJ.CODTRAB AND DIAS<>0 AND NOMBOL<>0 AND CERRADO=1"
    With Reporte
        .Reset
        .Connect = "DSN=" & VGL_SERVERREP & ";User=sa;PWD=;DSQ=planilla_pp"
        .ReportFileName = REGSISTEMA.REPORTES & "PLAN0085.RPT"
        .StoredProcParam(0) = " [##TMPVAC" & VGL_COMPUTER & "] "
        .Destination = crptToWindow
        .WindowShowPrintBtn = True
        .WindowShowRefreshBtn = True
        .WindowShowSearchBtn = True
        .WindowShowPrintSetupBtn = True
        .WindowState = crptMaximized
        .WindowTitle = .ReportFileName
        .Formulas(0) = "Empresa='" & REGSISTEMA.EMPRESA & "'"
        .Formulas(1) = "Titulo='" & UCase(ButtonMenu.Text) & " " & UCase(Option4.Caption) & "'"
        If .Status <> 2 Then .Action = 1
    End With
End If
 
End Sub


Private Sub XDATA_DBLCLICK()
    CMMODIFICAR_CLICK
End Sub

Private Sub XDATA_HEADCLICK(ByVal COLINDEX As Integer)
    RSTRABS.Sort = xData.Columns(COLINDEX).DataField
End Sub

Private Sub XMES_CHANGE()
    Select Case xMes.Tag
        Case 1
            OPTION1_CLICK
        Case 2
            OPTION2_Click
        Case 3
            OPTION3_Click
        Case 4
            OPTION4_Click
    End Select
End Sub

