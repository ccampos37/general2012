VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frAdminLiquid 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Panel de Administración de Liquidaciones"
   ClientHeight    =   4965
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9135
   Icon            =   "frAdminLiquid.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   9135
   Begin Crystal.CrystalReport reporte 
      Left            =   2115
      Top             =   1950
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Constancia"
      Height          =   390
      Left            =   7335
      TabIndex        =   10
      Top             =   675
      Width           =   1500
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Certificado"
      Height          =   390
      Left            =   7335
      TabIndex        =   9
      ToolTipText     =   "Certificado de Trabajo"
      Top             =   165
      Width           =   1500
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&Retiro CTS"
      Height          =   390
      Left            =   7335
      TabIndex        =   8
      Top             =   1185
      Width           =   1500
   End
   Begin VB.CommandButton cmCerrar 
      Cancel          =   -1  'True
      Caption         =   "&Cerrar"
      Height          =   360
      Left            =   7440
      TabIndex        =   7
      Top             =   4455
      Width           =   1485
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Anular Liq."
      Height          =   360
      Left            =   1990
      TabIndex        =   6
      Top             =   4395
      Width           =   1485
   End
   Begin MSComCtl2.DTPicker xFechaFin 
      Height          =   300
      Left            =   3585
      TabIndex        =   5
      Top             =   225
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   529
      _Version        =   393216
      Format          =   62455809
      CurrentDate     =   36864
   End
   Begin MSComCtl2.DTPicker xFechaIni 
      Height          =   300
      Left            =   885
      TabIndex        =   3
      Top             =   225
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   529
      _Version        =   393216
      Format          =   62455809
      CurrentDate     =   36864
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Liquidar"
      Height          =   360
      Left            =   210
      TabIndex        =   1
      Top             =   4395
      Width           =   1485
   End
   Begin MSDataGridLib.DataGrid xData 
      Height          =   3735
      Left            =   210
      TabIndex        =   0
      Top             =   615
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   6588
      _Version        =   393216
      AllowUpdate     =   -1  'True
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
      Caption         =   "Trabajadores Liquidados"
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "CodTrab"
         Caption         =   "Codigo"
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
         DataField       =   "Nombres"
         Caption         =   "Trabajador"
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
         DataField       =   "FechaIng"
         Caption         =   "Fecha Ing."
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
      BeginProperty Column03 
         DataField       =   "FechaCese"
         Caption         =   "Fecha Cese"
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
         MarqueeStyle    =   3
         BeginProperty Column00 
            ColumnWidth     =   794.835
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   3000.189
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1200.189
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hasta:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   2850
      TabIndex        =   4
      Top             =   285
      Width           =   465
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Desde:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   195
      TabIndex        =   2
      Top             =   285
      Width           =   510
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      Height          =   4725
      Left            =   120
      Top             =   120
      Width           =   7020
   End
End
Attribute VB_Name = "frAdminLiquid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RSLIQS As New ADODB.Recordset

Private Sub cmCerrar_Click()
    Unload Me
End Sub
Private Sub Command1_Click()
    frLiquidacion.Show 1
    REFRESCAR
End Sub

Private Sub Command2_Click()
    If RSLIQS.EOF Then Exit Sub
    If MsgBox("Seguro de Eliminar el Registro Seleccionado", vbYesNo + vbQuestion) = vbNo Then Exit Sub
    RegLiquida.CronoVac = DevuelveValor("SELECT CRONOVAC FROM LIQUIDACIONES WHERE CODIGO=" & RSLIQS!Codigo, DBSYSTEM)
    RegLiquida.CronoGrat = DevuelveValor("SELECT CRONOGRAT FROM LIQUIDACIONES WHERE CODIGO=" & RSLIQS!Codigo, DBSYSTEM)
    DBSYSTEM.Execute "DELETE FROM LIQUIDACIONES WHERE CODIGO=" & RSLIQS!Codigo
    DBSYSTEM.Execute "UPDATE TRABAJADORES SET SITUACIÓN='1', FECHACESE=NULL WHERE CODTRAB='" & RSLIQS!CODTRAB & "'"
    DBSYSTEM.Execute "DELETE FROM INGMOV2000 WHERE CODNOMBOL=" & RegLiquida.CronoVac & " AND CODTRAB='" & RSLIQS!CODTRAB & "' AND CONCEPTO='REMUVAC'"
    DBSYSTEM.Execute "DELETE FROM INGMOV2000 WHERE CODNOMBOL=" & RegLiquida.CronoGrat & " AND CODTRAB='" & RSLIQS!CODTRAB & "' AND CONCEPTO='REMUGRAT'"
    REFRESCAR
End Sub
Private Sub Command3_Click()
    If ExisteTablaAux("[##TMPVOUCHER" & VGL_COMPUTER & "]") Then DBAUXCOM.Execute "DROP TABLE [##TMPVOUCHER" & VGL_COMPUTER & "] "
    DBAUXCOM.Execute "CREATE TABLE [##TMPVOUCHER" & VGL_COMPUTER & "](CODTRAB VARCHAR(8))"
    
    DBSYSTEM.Execute "INSERT INTO  [##TMPVOUCHER" & VGL_COMPUTER & "]  VALUES ('" & RSLIQS!CODTRAB & "')"
    Dim RS1 As New ADODB.Recordset
    Set RS1 = New ADODB.Recordset
    RS1.Open "SELECT * FROM LIQUIDACIONES WHERE CODIGO=" & RSLIQS!Codigo, DBSYSTEM, adOpenStatic, adLockReadOnly
    Screen.MousePointer = 11
    CambiaPanelBD True
    With Reporte
        .Reset
        .WindowTitle = "PLAN0074.RPT - Resumen"
        .Connect = "DSN=" & VGL_SERVERREP & ";User=sa;PWD=;DSQ=planilla_pp"
        .ReportFileName = REGSISTEMA.REPORTES & "PLAN0074.RPT"
        .StoredProcParam(0) = REGSISTEMA.BASESQL
        .StoredProcParam(1) = VGL_COMPUTER
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .WindowShowPrintBtn = True
        .WindowShowRefreshBtn = True
        .WindowShowSearchBtn = True
        .WindowShowPrintSetupBtn = True
        
        .Formulas(0) = "XEMPRESA='" & REGSISTEMA.EMPRESA & "'"
        .Formulas(1) = "XRUC='" & REGSISTEMA.RUC & "'"
        .Formulas(2) = "XCTS=" & RS1!BASECTS
        .Formulas(3) = "XANNOCTS=" & RS1!a1
        .Formulas(4) = "XMESCTS=" & RS1!M1
        .Formulas(5) = "XDIASCTS=" & RS1!D1
        .Formulas(6) = "ANNO1=" & RS1!BASECTS * RS1!a1
        .Formulas(7) = "MES1=" & RS1!BASECTS / 12 * RS1!M1
        .Formulas(8) = "DIAS1=" & RS1!BASECTS / 12 / 30 * RS1!D1

        .Formulas(9) = "XVAC=" & RS1!BASEVAC
        .Formulas(10) = "XANNOVAC=" & RS1!A2
        .Formulas(11) = "XMESVAC=" & RS1!M2
        .Formulas(12) = "XDIASVAC=" & RS1!D2
        .Formulas(13) = "ANNO2=" & RS1!BASEVAC * RS1!A2
        .Formulas(14) = "MES2=" & RS1!BASEVAC / 12 * RS1!M2
        .Formulas(15) = "DIAS2=" & RS1!BASEVAC / 12 / 30 * RS1!D2

        .Formulas(16) = "XGRAT=" & RS1!BASEGRATI
        .Formulas(17) = "XANNOGRATI=" & RS1!A3
        .Formulas(18) = "XMESGRATI=" & RS1!M3
        .Formulas(19) = "XDIASGRATI=" & RS1!D3
        .Formulas(20) = "ANNO3=" & RS1!BASEGRATI * RS1!A3
        .Formulas(21) = "MES3=" & RS1!BASEGRATI / 12 * RS1!M3
        .Formulas(22) = "DIAS3=" & RS1!BASEGRATI / 12 / 30 * RS1!D3

        'SECCIÓN DE AFP
        .Formulas(23) = "XAFP='" & DevuelveValor("SELECT NOMBRE FROM AFPS WHERE CODAFP='" & RS1!CODAFP & "'", DBSYSTEM) & "'"
        .Formulas(24) = "XAFPVAC=" & RS1!AFP1
        .Formulas(25) = "XAFPGRAT=" & RS1!AFP2

        If .Status <> 2 Then .Action = 1
    End With
    Screen.MousePointer = 1
    CambiaPanelBD False
End Sub

Private Sub Command4_Click()
Dim RENUNCIA As String
Dim RS As New ADODB.Recordset
RS.Open "SELECT LIQUIDACIONES.CARGO, LIQUIDACIONES.FECHACESE, TRABAJADORES.APEPAT, TRABAJADORES.APEMAT, TRABAJADORES.NOMBRE AS NAME, TRABAJADORES.AREA, AREASTRAB.NOMBRE, LIQUIDACIONES.FECHAING  " & _
        " FROM AREASTRAB INNER JOIN (TRABAJADORES INNER JOIN LIQUIDACIONES ON TRABAJADORES.CODTRAB = LIQUIDACIONES.CODTRAB) ON AREASTRAB.CODCCOSTO = TRABAJADORES.AREA " & _
        " WHERE (((TRABAJADORES.CODTRAB)=[LIQUIDACIONES].[CODTRAB]) AND ((LIQUIDACIONES.CODTRAB)='" & RSLIQS!CODTRAB & "'));", DBSYSTEM, adOpenKeyset

If Not RS.RecordCount > 0 Then Exit Sub
    Screen.MousePointer = 11
    CambiaPanelBD True
    With Reporte
        .Reset
        .WindowTitle = "PLAN0080.RPT - CERTIFICADO DE TRABAJO"
        .ReportFileName = REGSISTEMA.REPORTES & "PLAN0080.RPT"
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .WindowShowPrintBtn = True
        .WindowShowRefreshBtn = True
        .WindowShowSearchBtn = True
        .WindowShowPrintSetupBtn = True
        RENUNCIA = InputBox("INGRESE EL MOTIVO DE CESE DEL TRABAJADOR", "INGRESO", "RENUNCIA VOLUNTARIA")
        DBSYSTEM.Execute "DELETE FROM MOTIVORENUNCIA WHERE CODTRAB='" & RSLIQS!CODTRAB & "'"
        DBSYSTEM.Execute "INSERT INTO MOTIVORENUNCIA VALUES ('" & RSLIQS!CODTRAB & "','" & RENUNCIA & "')"
        .Formulas(0) = "CARGO='" & RS!CARGO & "'"
        .Formulas(1) = "DEPENDE='" & RS!NOMBRE & "'"
        .Formulas(2) = "DESDE='" & Format(RS!FECHAING, "DD MMMM  YYYY") & "'"
        .Formulas(3) = "DISTRITO='LOS OLIVOS " & Format(Date, "DDDD, D MMMM  YYYY") & "'"
        .Formulas(4) = "EMPRESA='CAMTEX S.A.'"
        .Formulas(5) = "EMPTRAB='" & REGSISTEMA.EMPRESA & "'"
        .Formulas(6) = "FIRMA='DPTO. RECURSOS HUMANOS'"
        .Formulas(8) = "HASTA='" & Format(RS!FECHACESE, "DD MMMM  YYYY") & "'"
        .Formulas(9) = "MOTIVO='" & RENUNCIA & "'"
        .Formulas(10) = "NOMBRE='" & RS!ApePat & " " & RS!ApeMat & ", " & RS!Name & "'"

        If .Status <> 2 Then .Action = 1
    End With
    Screen.MousePointer = 1
    CambiaPanelBD False
End Sub

Private Sub Command5_Click()
Dim RS As New ADODB.Recordset
Dim RSAUX As New ADODB.Recordset
Dim EXISTE As Boolean
EXISTE = False
RS.Open "SELECT TRABAJADORES.APEPAT, TRABAJADORES.APEMAT, TRABAJADORES.NOMBRE AS NAME, TRABAJADORES.AREA, AREASTRAB.NOMBRE, BANCOS.NOMBRE AS BANCO, TRABAJADORES.CTABANCO, TRABAJADORES.CODTRAB " & _
        " FROM BANCOS INNER JOIN (AREASTRAB INNER JOIN TRABAJADORES ON AREASTRAB.CODCCOSTO = TRABAJADORES.AREA) ON BANCOS.CODBANCO = TRABAJADORES.BANCO " & _
        " WHERE (((TRABAJADORES.CODTRAB)='" & RSLIQS!CODTRAB & "'))", DBSYSTEM, adOpenKeyset
If ExisteTabla("MOTIVORENUNCIA") Then
    RSAUX.Open "SELECT * FROM MOTIVORENUNCIA WHERE CODTRAB='" & RSLIQS!CODTRAB & "'", DBSYSTEM, adOpenKeyset, adLockOptimistic
    EXISTE = True
End If
If Not RS.RecordCount > 0 Then Exit Sub
    Screen.MousePointer = 11
    CambiaPanelBD True
    With Reporte
        .Reset
        .WindowTitle = "PLAN0081.RPT - RETIRO CTS"
        .ReportFileName = REGSISTEMA.REPORTES & "PLAN0081.RPT"
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .WindowShowPrintBtn = True
        .WindowShowRefreshBtn = True
        .WindowShowSearchBtn = True
        .WindowShowPrintSetupBtn = True
        .Formulas(0) = "SNR='" & RS!BANCO & "'"
        .Formulas(1) = "F1='" & Format(Date, "DD MMMM  YYYY") & "'"
        .Formulas(2) = "DISTRITO='LOS OLIVOS " & Format(Date, "DDDD, D MMMM  YYYY") & "'"
        .Formulas(3) = "EMPRESA='CAMTEX S.A.'"
        .Formulas(4) = "FIRMA='ATENTAMENTE,'"
        .Formulas(5) = "CUENTA='" & RS!CTABANCO & "'"
        .Formulas(6) = "MOTIVO='" & IIf(EXISTE, IIf(RSAUX.RecordCount > 0, RSAUX.Fields(1), "RENUNCIA VOLUNTARIA"), " ") & "'"
        .Formulas(7) = "EMPLEADO='" & RS!ApePat & " " & RS!ApeMat & ", " & RS!Name & "'"
        If .Status <> 2 Then .Action = 1
    End With
    Screen.MousePointer = 1
    CambiaPanelBD False
End Sub
Private Sub Form_Load()
    xFechaFin.Value = Date
    xFechaIni.Value = xFechaFin.Value
    xFechaIni.Day = 1
    xFechaIni.Month = 1
    REFRESCAR
End Sub
Public Sub REFRESCAR()
    Set RSLIQS = Nothing
    RSLIQS.Open "SELECT A.CODTRAB,CODIGO, NOMBRES, A.FECHAING, A.FECHACESE FROM LIQUIDACIONES A, VWTRABAJ B WHERE A.CODTRAB=B.CODTRAB AND (A.FECHACESE BETWEEN " & DateSQL(xFechaIni.Value) & " AND " & DateSQL(xFechaFin.Value) & ")", DBSYSTEM, adOpenStatic, adLockReadOnly
    Set xData.DataSource = RSLIQS
End Sub
Private Sub FORM_UNLOAD(CANCEL As Integer)
    Set RSLIQS = Nothing
End Sub
Private Sub XDATA_HEADCLICK(ByVal COLINDEX As Integer)
    RSLIQS.Sort = xData.Columns(COLINDEX).DataField
End Sub

