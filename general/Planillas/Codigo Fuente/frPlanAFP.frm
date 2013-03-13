VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C5FF36B7-A67E-11D3-9D9D-E6F193F7F854}#9.0#0"; "ApliCTxt.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frPlanAFP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Planilla de Pago de AFP"
   ClientHeight    =   6390
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6825
   Icon            =   "frPlanAFP.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   6825
   Begin Crystal.CrystalReport RptAFP 
      Left            =   2655
      Top             =   2820
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmExportar 
      Caption         =   "&Cerrar"
      Height          =   330
      Left            =   5115
      TabIndex        =   40
      Top             =   5970
      Width           =   1305
   End
   Begin VB.CommandButton cmImprimir 
      Caption         =   "&Imprimir"
      Height          =   330
      Left            =   3675
      TabIndex        =   39
      Top             =   5970
      Width           =   1305
   End
   Begin VB.Frame Frame4 
      Caption         =   "Tipo de Pago"
      Height          =   2190
      Left            =   3405
      TabIndex        =   28
      Top             =   3615
      Width           =   3315
      Begin AplisetControlText.Aplitext xLiqNum 
         Height          =   285
         Left            =   1485
         TabIndex        =   38
         ToolTipText     =   "Ingrese número de Liquidación de Cobranza Judiacial"
         Top             =   1770
         Visible         =   0   'False
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   503
         MaxLength       =   15
         Text            =   ""
      End
      Begin VB.OptionButton fTipoPago 
         Caption         =   "Liquidación de Cobranza"
         Height          =   210
         Index           =   3
         Left            =   165
         TabIndex        =   36
         Top             =   1560
         Width           =   2055
      End
      Begin AplisetControlText.Aplitext xRegNum 
         Height          =   285
         Left            =   1485
         TabIndex        =   34
         ToolTipText     =   "Ingrese número de planilla de regularización"
         Top             =   1245
         Visible         =   0   'False
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   503
         MaxLength       =   15
         Text            =   ""
      End
      Begin VB.OptionButton fTipoPago 
         Caption         =   "Regularización de Planilla"
         Height          =   225
         Index           =   2
         Left            =   165
         TabIndex        =   33
         Top             =   1065
         Width           =   2175
      End
      Begin AplisetControlText.Aplitext xInteres 
         Height          =   285
         Left            =   1935
         TabIndex        =   32
         ToolTipText     =   "Ingrese interés moratorio"
         Top             =   735
         Visible         =   0   'False
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   503
         MaxLength       =   15
         Text            =   ""
      End
      Begin VB.OptionButton fTipoPago 
         Caption         =   "Extemporáneo"
         Height          =   210
         Index           =   1
         Left            =   165
         TabIndex        =   30
         ToolTipText     =   "Para aquellos pagos fuera del plazo dado por la SAFP"
         Top             =   540
         Width           =   2730
      End
      Begin VB.OptionButton fTipoPago 
         Caption         =   "Normal"
         Height          =   195
         Index           =   0
         Left            =   165
         TabIndex        =   29
         ToolTipText     =   "Pago normal, la cual corresponde con la fecha de pago y un cornograma de la SAFP"
         Top             =   285
         UseMaskColor    =   -1  'True
         Value           =   -1  'True
         Width           =   1185
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Número"
         Height          =   195
         Left            =   450
         TabIndex        =   37
         Top             =   1815
         Width           =   555
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Número"
         Height          =   195
         Left            =   450
         TabIndex        =   35
         Top             =   1290
         Width           =   555
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Interés Moratorio:"
         Height          =   195
         Left            =   450
         TabIndex        =   31
         Top             =   780
         Width           =   1230
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Forma de Pago por Retenciones (AFP)"
      Height          =   1290
      Left            =   120
      TabIndex        =   21
      Top             =   5010
      Width           =   3180
      Begin AplisetControlText.Aplitext xf2Banco 
         Height          =   285
         Left            =   1140
         TabIndex        =   27
         ToolTipText     =   "Banco del Cheque para pagar a la AFP(Comisiones)"
         Top             =   885
         Visible         =   0   'False
         Width           =   1890
         _ExtentX        =   3334
         _ExtentY        =   503
         MaxLength       =   25
         Text            =   ""
      End
      Begin AplisetControlText.Aplitext xf2Cheque 
         Height          =   285
         Left            =   1140
         TabIndex        =   25
         ToolTipText     =   "Número de Cheque para pagar a la AFP (Comisiones)"
         Top             =   585
         Visible         =   0   'False
         Width           =   1890
         _ExtentX        =   3334
         _ExtentY        =   503
         MaxLength       =   20
         Text            =   ""
      End
      Begin VB.OptionButton f2Tipo 
         Caption         =   "Cheque"
         Height          =   195
         Index           =   1
         Left            =   1530
         TabIndex        =   23
         ToolTipText     =   "Pago realizado con cheque bancario. Solo cuando está dentro del cronograma de la SAFP"
         Top             =   315
         Width           =   945
      End
      Begin VB.OptionButton f2Tipo 
         Caption         =   "Efectivo"
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   22
         ToolTipText     =   "Para realizado con dinero en efectivo"
         Top             =   315
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Banco"
         Height          =   195
         Left            =   195
         TabIndex        =   26
         Top             =   930
         Width           =   465
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Cheque N°"
         Height          =   195
         Left            =   180
         TabIndex        =   24
         Top             =   630
         Width           =   780
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Forma de Pago al Fondo de Pensiones"
      Height          =   1290
      Left            =   120
      TabIndex        =   14
      Top             =   3615
      Width           =   3180
      Begin AplisetControlText.Aplitext xf1Banco 
         Height          =   285
         Left            =   1170
         TabIndex        =   20
         ToolTipText     =   "Banco del Cheque para pagar al fondo de pensiones"
         Top             =   870
         Visible         =   0   'False
         Width           =   1890
         _ExtentX        =   3334
         _ExtentY        =   503
         MaxLength       =   25
         Text            =   ""
      End
      Begin AplisetControlText.Aplitext xf1Cheque 
         Height          =   285
         Left            =   1170
         TabIndex        =   18
         ToolTipText     =   "Número de Cheque para pagar al Fondo de Pensiones"
         Top             =   570
         Visible         =   0   'False
         Width           =   1890
         _ExtentX        =   3334
         _ExtentY        =   503
         MaxLength       =   20
         Text            =   ""
      End
      Begin VB.OptionButton f1Tipo 
         Caption         =   "Cheque"
         Height          =   210
         Index           =   1
         Left            =   1530
         TabIndex        =   16
         ToolTipText     =   "Pago realizado con cheque bancario. Solo cuando está dentro del cronograma de la SAFP"
         Top             =   315
         Width           =   1305
      End
      Begin VB.OptionButton f1Tipo 
         Caption         =   "Efectivo"
         Height          =   210
         Index           =   0
         Left            =   150
         TabIndex        =   15
         ToolTipText     =   "Para realizado con dinero en efectivo"
         Top             =   315
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Banco"
         Height          =   195
         Left            =   165
         TabIndex        =   19
         Top             =   915
         Width           =   465
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Cheque N°"
         Height          =   195
         Left            =   165
         TabIndex        =   17
         Top             =   600
         Width           =   780
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos del Empleador"
      Height          =   1560
      Left            =   120
      TabIndex        =   8
      Top             =   1980
      Width           =   6600
      Begin MSComCtl2.DTPicker xFecha 
         Height          =   285
         Left            =   4950
         TabIndex        =   3
         Top             =   525
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   503
         _Version        =   393216
         Format          =   60424193
         CurrentDate     =   36672
      End
      Begin AplisetControlText.Aplitext xResponsable 
         Height          =   285
         Left            =   1395
         TabIndex        =   1
         ToolTipText     =   "Persona/Funcionario responsable de la elaboración de AFP"
         Top             =   225
         Width           =   3480
         _ExtentX        =   6138
         _ExtentY        =   503
         MaxLength       =   25
         Text            =   ""
      End
      Begin AplisetControlText.Aplitext xDepartamento 
         Height          =   285
         Left            =   1395
         TabIndex        =   2
         ToolTipText     =   "Departamento/Area/Sección responsable de AFP dentro de la empresa"
         Top             =   525
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   503
         MaxLength       =   25
         Text            =   ""
      End
      Begin VB.CommandButton xDatosEmp 
         Caption         =   "Datos Empleador"
         Height          =   285
         Left            =   4950
         TabIndex        =   7
         ToolTipText     =   "Ver los datos generales de la empresa"
         Top             =   1155
         Width           =   1485
      End
      Begin AplisetControlText.Aplitext xBanco 
         Height          =   285
         Left            =   1395
         TabIndex        =   6
         ToolTipText     =   "Institución bancaria de referencia para la AFP"
         Top             =   1125
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   503
         MaxLength       =   20
         Text            =   ""
      End
      Begin AplisetControlText.Aplitext xTipoCta 
         Height          =   285
         Left            =   4950
         TabIndex        =   5
         ToolTipText     =   "Tipo de Cuenta Bancaria de la Empresa: Ej. Cta. Cte./Cta. Ahorro"
         Top             =   825
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   503
         MaxLength       =   15
         Text            =   ""
      End
      Begin AplisetControlText.Aplitext xCtaBanco 
         Height          =   285
         Left            =   1395
         TabIndex        =   4
         ToolTipText     =   "Cuenta Bancaria para referencia de la AFP"
         Top             =   825
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   503
         MaxLength       =   20
         Text            =   ""
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Pago"
         Height          =   195
         Left            =   3945
         TabIndex        =   41
         Top             =   555
         Width           =   870
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Inst. Bancaria"
         Height          =   195
         Left            =   150
         TabIndex        =   13
         Top             =   1170
         Width           =   975
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Cuenta"
         Height          =   195
         Left            =   3945
         TabIndex        =   12
         Top             =   870
         Width           =   870
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta Bancaria"
         Height          =   195
         Left            =   150
         TabIndex        =   11
         Top             =   870
         Width           =   1185
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Departamento"
         Height          =   195
         Left            =   165
         TabIndex        =   10
         Top             =   570
         Width           =   1005
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Responsable"
         Height          =   195
         Left            =   150
         TabIndex        =   9
         Top             =   270
         Width           =   930
      End
   End
   Begin MSDataGridLib.DataGrid dgAFPs 
      Height          =   1785
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6585
      _ExtentX        =   11615
      _ExtentY        =   3149
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
      Caption         =   "Periodo de Devengue: 04/2000"
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
            LCID            =   2058
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
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   4
         AllowRowSizing  =   0   'False
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frPlanAFP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public RSRESAFP As New ADODB.Recordset 'RESUMEN DE AFP

Private Sub CMEXPORTAR_CLICK()
    Unload Me
End Sub

Private Sub CMIMPRIMIR_CLICK()
    If RSRESAFP!NUMTRAB = 0 Or IsNull(RSRESAFP!NUMTRAB) Then
        MsgBox "NO EXISTEN DATOS POR IMPRIMIR", vbCritical
        Exit Sub
    End If
    If RSRESAFP!NUMPLANILLA = "" Or IsNull(RSRESAFP!NUMPLANILLA) Then
        MsgBox "FALTA CONSIGNAR EL NÚMERO DE PLANILLA", vbCritical
        Exit Sub
    End If
    cmImprimir.Tag = RSRESAFP!NUMTRAB
    FrPagReg.Show 1
End Sub

Private Sub F1TIPO_Click(INDEX As Integer)
    If INDEX = 0 Then
        xf1Cheque.Visible = False
        xf1Banco.Visible = False
    Else
        xf1Cheque.Visible = True
        xf1Banco.Visible = True
        xf1Cheque.SetFocus
    End If
End Sub

Private Sub F2TIPO_Click(INDEX As Integer)
    If INDEX = 0 Then
        xf2Cheque.Visible = False
        xf2Banco.Visible = False
    Else
        xf2Cheque.Visible = True
        xf2Banco.Visible = True
        xf2Cheque.SetFocus
    End If
End Sub

Private Sub Form_Load()
    CARGARPLAN CDate(VPTAREA)
    xFecha.MinDate = CDate(VPTAREA)
    xResponsable.Text = GetSetting(App.CompanyName, "AFP", "RESPON", "")
    xDepartamento.Text = GetSetting(App.CompanyName, "AFP", "DEPARTAM", "")
    xCtaBanco.Text = GetSetting(App.CompanyName, "AFP", "CTACTE", "")
    xBanco.Text = GetSetting(App.CompanyName, "AFP", "BANCO", "")
    xTipoCta.Text = GetSetting(App.CompanyName, "AFP", "TIPOCTA", "")
End Sub

Public Sub CARGARPLAN(ByVal VARMES As Date)
    If ExisteTablaAux("[##_TMPAFP" & VGL_COMPUTER & "]") Then DBSYSTEM.Execute "DROP TABLE  [##_TMPAFP" & VGL_COMPUTER & "] "
    DBSYSTEM.Execute "CREATE TABLE  [##_TMPAFP" & VGL_COMPUTER & "]  (CODIGO VARCHAR(2), NOMAFP VARCHAR(25), NUMTRAB INT, NUMPLANILLA VARCHAR(15), TOTALFP  Numeric(20,2) , TOTALAFP  Numeric(20,2) , TOTAL  Numeric(20,2) )"
    RSRESAFP.Open "SELECT * FROM AFPS ORDER BY NOMBRE", DBSYSTEM, adOpenStatic, adLockOptimistic
    Do While Not RSRESAFP.EOF
        If RSRESAFP!CODAFP <> "ON" Then DBSYSTEM.Execute "INSERT INTO  [##_TMPAFP" & VGL_COMPUTER & "]  (CODIGO, NOMAFP) VALUES ('" & RSRESAFP!CODAFP & "','" & RSRESAFP!NOMBRE & "')"
        RSRESAFP.MoveNext
    Loop
    'RSRESAFP.CLOSE
    'RSRESAFP.OPEN "SELECT COUNT(CODTRAB) AS TOTAL, FONDOPENS FROM " & REGSISTEMA.TABLAPLAN & " WHERE MES=" & DATESQL(VARMES) & " AND FONDOPENS<>'ON' GROUP BY FONDOPENS", DBSYSTEM, ADOPENSTATIC
    RSRESAFP.MoveFirst
    Dim RSAX As New ADODB.Recordset
    'CALCULO DEL NUMERO DE TRABAJADORES
    Do While Not RSRESAFP.EOF
        Set RSAX = Nothing
        RSAX.Open "SELECT DISTINCT CODTRAB FROM " & REGSISTEMA.TABLAPLAN & " WHERE MES=" & DateSQL(VARMES) & " AND FONDOPENS='" & RSRESAFP!CODAFP & "' AND CODTRAB IN (SELECT CODTRAB FROM TRABAJADORES WHERE " & IIf(VPNUMTMP = 1, "AREA", "CCOSTO") & " IN " & VPTRASPRM & " AND FONDOPENS='" & RSRESAFP!CODAFP & "')", DBSYSTEM, adOpenStatic, adLockReadOnly
        DBSYSTEM.Execute "UPDATE  [##_TMPAFP" & VGL_COMPUTER & "]  SET NUMTRAB=" & RSAX.RecordCount & " WHERE CODIGO='" & RSRESAFP!CODAFP & "'"
        RSRESAFP.MoveNext
    Loop
    dgAFPs.Caption = "PERIODO DE DEVENGUE: " & Format(Month(VARMES), "00") & "/" & Year(VARMES)
    Set RSAX = Nothing
    RSRESAFP.Close
    If CARGAPLAFP(VARMES) = False Then Exit Sub
    
    RSRESAFP.Open "SELECT SUM(REMUASEG) AS T1, SUM(APOROBLI) AS T2, SUM(SEGUROS) AS T3, SUM(COMISION) AS T4, CODAFP FROM  [##_TMPPLANAFP" & VGL_COMPUTER & "]  GROUP BY CODAFP", DBSYSTEM, adOpenStatic
    Do While Not RSRESAFP.EOF
        DBSYSTEM.Execute "UPDATE  [##_TMPAFP" & VGL_COMPUTER & "]  SET TOTALFP=" & RSRESAFP!T2 & " WHERE CODIGO='" & RSRESAFP!CODAFP & "'"
        DBSYSTEM.Execute "UPDATE  [##_TMPAFP" & VGL_COMPUTER & "]  SET TOTALAFP=" & RSRESAFP!T3 & "+" & RSRESAFP!T4 & " WHERE CODIGO='" & RSRESAFP!CODAFP & "'"
        RSRESAFP.MoveNext
    Loop
    DBSYSTEM.Execute "UPDATE  [##_TMPAFP" & VGL_COMPUTER & "]  SET TOTAL=TOTALFP+TOTALAFP"
    RSRESAFP.Close
    RSRESAFP.Open " [##_TMPAFP" & VGL_COMPUTER & "] ", DBSYSTEM, adOpenKeyset, adLockOptimistic
    Set dgAFPs.DataSource = RSRESAFP
    With dgAFPs
        .Columns(0).Width = 540.2835
        .Columns(1).Width = 1065.26
        .Columns(2).Width = 629.8583
        .Columns(3).Width = 1230.236
        .Columns(4).Width = 915.0237
        .Columns(5).Width = 915.0237
        .Columns(6).Width = 915.0237
        .Columns("CODIGO").Locked = True
        .Columns("NOMAFP").Locked = True
        .Columns("NUMTRAB").Locked = True
        .Columns("TOTALFP").Locked = True
        .Columns("TOTALAFP").Locked = True
        .Columns("TOTAL").Locked = True
        .Columns("TOTALFP").NumberFormat = "##,##0.00 "
        .Columns("TOTALAFP").NumberFormat = "##,##0.00 "
        .Columns("TOTAL").NumberFormat = "##,##0.00 "
        .Columns("TOTALFP").Alignment = dbgRight
        .Columns("TOTALAFP").Alignment = dbgRight
        .Columns("TOTAL").Alignment = dbgRight
        .Columns("NUMTRAB").Alignment = dbgCenter
    End With
End Sub

Private Sub FORM_UNLOAD(CANCEL As Integer)
    Set RSRESAFP = Nothing
End Sub

Private Sub FTIPOPAGO_Click(INDEX As Integer)
    xInteres.Visible = False
    xRegNum.Visible = False
    xLiqNum.Visible = False
    Select Case INDEX
        Case 1: xInteres.Visible = True: xInteres.SetFocus
        Case 2: xRegNum.Visible = True: xRegNum.SetFocus
        Case 3: xLiqNum.Visible = True: xLiqNum.SetFocus
    End Select
End Sub

Private Sub XBANCO_DblClick()
    Dim RSBAN As New ADODB.Recordset
    RSBAN.Open "SELECT CODBANCO, NOMBRE FROM BANCOS ORDER BY NOMBRE", DBSYSTEM, adOpenStatic
    If RSBAN.RecordCount = 0 Then
        MsgBox "NO SE PUEDE DESPLEGAR LA AYUDA, PUES NO EXISTEN BANCOS REGISTRADOS", vbCritical
        Exit Sub
    End If
    frmComun.CONECTAR RSBAN
    frmComun.Show 1
    If VGUTIL(1) <> "" Then
        xBanco.Text = VGUTIL(2)
    End If
    Set RSBAN = Nothing
End Sub

Private Sub XDATOSEMP_Click()
    frDatos.Show 1
End Sub

Function CARGAPLAFP(ByVal VALMES As Date) As Boolean
CARGAPLAFP = False
    Dim RSAUX As New ADODB.Recordset
    Dim VALAFP(5) As String, NOMTABBOL As String, NOMTABMOV As String, CADENAIN As String
    RSAUX.Open "EMPRESA", DBSYSTEM, adOpenStatic
    With RSAUX
        If !AFPREMU = "" Then
            MsgBox "NO SE HA CONFIGURADO LA ASIGNACIÓN SOBRE REMUNERACIÓN ASEGURABLE", vbCritical
            Set RSAUX = Nothing
            Exit Function
        End If
        If !AFPAPOR = "" Then
            MsgBox "NO SE HA CONFIGURADO LA APORTACIÓN OBLIGATORIA", vbCritical
            Set RSAUX = Nothing
            Exit Function
        End If
        If !AFPSEG = "" Then
            MsgBox "NO SE HA CONFIGURADO EL SEGURO DE VIDA DE AFP", vbCritical
            Set RSAUX = Nothing
            Exit Function
        End If
        If !AFPCOMI = "" Then
            MsgBox "NO SE HA CONFIGURADO LA COMISIÓN % SOBRE REMUNERACIÓN ASEGURABLE", vbCritical
            Set RSAUX = Nothing
            Exit Function
        End If
        If !AFPPLAN = "" Then
            MsgBox "NO SE HA CONFIGURADO EL RUBRO PARA REALIZAR LA CONSISTENCIA DE VALORES ENTRE PLANILLA DE AFP Y PLANILLA DE REMUNERACIONES", vbCritical
            Set RSAUX = Nothing
            Exit Function
        End If
        VALAFP(0) = !AFPREMU
        VALAFP(1) = !AFPAPOR
        VALAFP(2) = !AFPSEG
        VALAFP(3) = !AFPCOMI
        VALAFP(4) = !AFPPLAN
        CADENAIN = "'" & VALAFP(1) & "','" & VALAFP(2) & "','" & VALAFP(3) & "'"
    End With
    If ExisteTablaAux(" [##_TMPPLANAFP" & VGL_COMPUTER & "] ") Then DBSYSTEM.Execute "DROP TABLE  [##_TMPPLANAFP" & VGL_COMPUTER & "] "
    If ExisteTablaAux(" [##_TMPPLANAFP1" & VGL_COMPUTER & "] ") Then DBSYSTEM.Execute "DROP TABLE  [##_TMPPLANAFP1" & VGL_COMPUTER & "] "
    If ExisteTablaAux(" [##_TMPPLANAFP2" & VGL_COMPUTER & "] ") Then DBSYSTEM.Execute "DROP TABLE  [##_TMPPLANAFP2" & VGL_COMPUTER & "] "
    If ExisteTablaAux("[##_TMPPLANAFPREST" & VGL_COMPUTER & "] ") Then DBSYSTEM.Execute "DROP TABLE  [##_TMPPLANAFPREST" & VGL_COMPUTER & "] "
    DBSYSTEM.Execute "CREATE TABLE  [##_TMPPLANAFP" & VGL_COMPUTER & "]  (INUMBOL INT, CODAFP VARCHAR(2), CODTRAB VARCHAR(8), CUSPP VARCHAR(12), APEPAT VARCHAR(20), APEMAT VARCHAR(20), NOMBRES VARCHAR(20), TIPO VARCHAR(1), FECHAMOV DATETIME, REMUASEG  Numeric(20,2) , APOROBLI  Numeric(20,2) , APORVOLT  Numeric(20,2) , APORVOLE  Numeric(20,2) , APOREMP  Numeric(20,2) , SEGUROS  Numeric(20,2) , COMISION  Numeric(20,2) )"
    RSAUX.Close
    Dim RSAFP As New ADODB.Recordset
    Dim XINUMERO As Long
    XINUMERO = 0
    RSAFP.Open " [##_TMPPLANAFP" & VGL_COMPUTER & "] ", DBSYSTEM, adOpenKeyset, adLockOptimistic
    NOMTABBOL = "BOL" & Format(Month(VALMES), "00") & Year(VALMES)
    NOMTABMOV = "MOV" & Format(Month(VALMES), "00") & Year(VALMES)

'ERROR AL EJECUTAR LA CONSULTA
    RSAUX.Open "SELECT BOL.CODTRAB, BOL.CODAFP, CUSPP, APEPAT, APEMAT, NOMBRE, FECHAING, FECHACESE, MONTO, CONCEPTO, BOL.INUMBOL, BOL.SUMAAFP FROM TRABAJADORES," & NOMTABBOL & " BOL ," & NOMTABMOV & " MOV WHERE BOL.CODTRAB=TRABAJADORES.CODTRAB AND BOL.INUMBOL=MOV.INUMBOL AND BOL.CODAFP<>'ON' AND " & IIf(VPNUMTMP = 1, "AREA", "TRABAJADORES.CCOSTO") & " IN " & VPTRASPRM & " AND MOV.CONCEPTO IN (" & CADENAIN & ") AND BOL.CODNOMBOL IN (SELECT CODIGO FROM NOMBOL WHERE MES=" & DateSQL(VALMES) & ")  ORDER BY BOL.CODTRAB,BOL.INUMBOL,CONCEPTO ", DBSYSTEM, adOpenStatic
'MODIFICAR PRONTO
    With RSAUX
    Do While Not RSAUX.EOF
        If RSAFP.RecordCount > 0 Then
            RSAFP.MoveFirst
            RSAFP.FIND "CODTRAB='" & RSAUX!CODTRAB & "'"
        End If
        If RSAFP.EOF Then
            RSAFP.AddNew
            RSAFP!INUMBOL = RSAUX!INUMBOL
            RSAFP!CODAFP = RSAUX!CODAFP
            RSAFP!CODTRAB = RSAUX!CODTRAB
            RSAFP!CUSPP = RSAUX!CUSPP
            RSAFP!ApePat = RSAUX!ApePat
            RSAFP!ApeMat = RSAUX!ApeMat
            RSAFP!NOMBRES = RSAUX!NOMBRE
            RSAFP!TIPO = "0"
            If Format(Month(RSAUX!FECHAING), "00") & Year(RSAUX!FECHAING) = Format(Month(VALMES), "00") & Year(VALMES) Then
                RSAFP!TIPO = "1"
                RSAFP!FECHAMOV = FechS(RSAUX!FECHAING, Adof)
            End If
            If Not IsNull(RSAUX!FECHACESE) Then
                If Format(Month(RSAUX!FECHACESE), "00") & Year(RSAUX!FECHACESE) = Format(Month(VALMES), "00") & Year(VALMES) Then
                    RSAFP!TIPO = "2"
                    RSAFP!FECHAMOV = FechS(RSAUX!FECHACESE, Adof)
                End If
            End If
            RSAFP!REMUASEG = 0
            RSAFP!APOROBLI = 0
            RSAFP!SEGUROS = 0
            RSAFP!COMISION = 0
        End If
        If XINUMERO <> RSAUX!INUMBOL Then
            RSAFP!REMUASEG = RSAFP!REMUASEG + RSAUX!SUMAAFP
            XINUMERO = RSAUX!INUMBOL
        End If
        Select Case RSAUX!CONCEPTO
            Case VALAFP(1): RSAFP!APOROBLI = RSAFP!APOROBLI + RSAUX!MONTO
            Case VALAFP(2): RSAFP!SEGUROS = RSAFP!SEGUROS + RSAUX!MONTO
            Case VALAFP(3): RSAFP!COMISION = RSAFP!COMISION + RSAUX!MONTO
            Case Else
                Debug.Print "ERROR: " & RSAUX!CONCEPTO
        End Select
        RSAFP.Update
        RSAUX.MoveNext
    Loop
    End With
    Set RSAFP = Nothing
    Set RSAUX = Nothing
    DBSYSTEM.Execute "UPDATE  [##_TMPPLANAFP" & VGL_COMPUTER & "]  SET APORVOLT=0,APORVOLE=0,APOREMP=0"
    CARGAPLAFP = True
End Function

Private Sub XF1BANCO_DblClick()
    Dim RSBAN As New ADODB.Recordset
    RSBAN.Open "SELECT CODBANCO, NOMBRE FROM BANCOS ORDER BY NOMBRE", DBSYSTEM, adOpenStatic
    If RSBAN.RecordCount = 0 Then
        MsgBox "NO SE PUEDE DESPLEGAR LA AYUDA, PUES NO EXISTEN BANCOS REGISTRADOS", vbCritical
        Exit Sub
    End If
    frmComun.CONECTAR RSBAN
    frmComun.Show 1
    If VGUTIL(1) <> "" Then
        xf1Banco.Text = VGUTIL(2)
    End If
    Set RSBAN = Nothing
End Sub

Private Sub XF2BANCO_DblClick()
    Dim RSBAN As New ADODB.Recordset
    RSBAN.Open "SELECT CODBANCO, NOMBRE FROM BANCOS ORDER BY NOMBRE", DBSYSTEM, adOpenStatic
    If RSBAN.RecordCount = 0 Then
        MsgBox "NO SE PUEDE DESPLEGAR LA AYUDA, PUES NO EXISTEN BANCOS REGISTRADOS", vbCritical
        Exit Sub
    End If
    frmComun.CONECTAR RSBAN
    frmComun.Show 1
    If VGUTIL(1) <> "" Then
        xf2Banco.Text = VGUTIL(2)
    End If
    Set RSBAN = Nothing
End Sub

