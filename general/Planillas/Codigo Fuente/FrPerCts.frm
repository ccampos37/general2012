VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form FrPerCts 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Personalizar Certificado de CTS"
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4890
   Icon            =   "FrPerCts.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   4890
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   1875
      Left            =   45
      TabIndex        =   4
      Top             =   -60
      Width           =   4815
      Begin MSComCtl2.DTPicker DTPEmi 
         Height          =   300
         Left            =   1890
         TabIndex        =   6
         Top             =   600
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   529
         _Version        =   393216
         Format          =   62062593
         CurrentDate     =   37664
      End
      Begin MSComCtl2.DTPicker DTPDepo 
         Height          =   300
         Left            =   1905
         TabIndex        =   5
         Top             =   225
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   529
         _Version        =   393216
         Format          =   62062593
         CurrentDate     =   37664
      End
      Begin Crystal.CrystalReport Reporte 
         Left            =   1965
         Top             =   1140
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Cancelar"
         Height          =   330
         Left            =   3075
         TabIndex        =   3
         Top             =   1410
         Width           =   1665
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Aceptar"
         Height          =   330
         Left            =   3075
         TabIndex        =   2
         Top             =   1020
         Width           =   1665
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha de Emisión :"
         Height          =   285
         Left            =   180
         TabIndex        =   0
         Top             =   630
         Width           =   1485
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   420
         Picture         =   "FrPerCts.frx":0442
         Top             =   1140
         Width           =   480
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha de Depósito :"
         Height          =   285
         Left            =   195
         TabIndex        =   1
         Top             =   270
         Width           =   1530
      End
   End
End
Attribute VB_Name = "FrPerCts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim xPeriodo As String
Dim FECIN As Date
Dim FECFI As Date
Private Sub Command1_Click()
Dim RS As New ADODB.Recordset
Dim RSAUX As New ADODB.Recordset
Dim Codigo As String
Dim TOTAL As Single
Dim X As Integer
'If Not Len(Trim(Xcargo.Text)) > 0 Then
'    MsgBox "INGRESE EL DECRETO LEY O DECRETO DE EMERGENCIA REFERENTE A ESTE CERTIFICADO", vbCritical
'    Xcargo.SetFocus
'    Exit Sub
'End If
    Screen.MousePointer = 11
    Codigo = ""
    DBSYSTEM.Execute "ALTER TABLE ##TMPCTS4 ADD TOTAL  Numeric(20,2) "
    DBSYSTEM.Execute "ALTER TABLE ##TMPCTS4 ADD MONEDA varchar(10)"
    RS.Open "SELECT * FROM ##TMPCTS4 ORDER BY CODTRAB", DBSYSTEM, adOpenKeyset, adLockOptimistic
    If RS.RecordCount > 0 Then
        Codigo = RS!CODTRAB
        While Not RS.EOF
            If Codigo <> RS!CODTRAB Then
                DBSYSTEM.Execute "UPDATE ##TMPCTS4 SET TOTAL=" & Round(DevuelveValor("SELECT SUM(IMPORTECTS) FROM  [##TMPCTS3" & VGL_COMPUTER & "]  WHERE CODTRAB='" & Codigo & "'", DBSYSTEM), 2) & " WHERE CODTRAB='" & Codigo & "'"
                TOTAL = 0
                Codigo = RS!CODTRAB
            End If
            Set RSAUX = New ADODB.Recordset
            RSAUX.Open "SELECT TRABAJADORES.DOCIDEN, TRABAJADORES.APEMAT, TRABAJADORES.APEPAT, TRABAJADORES.NOMBRE AS EMPLEADO, TRABAJADORES.BANCO, TRABAJADORES.FECHAING, TRABAJADORES.CTACTS, BANCOS.NOMBRE,MONE=ISNULL(TRABAJADORES.MON,'01') FROM TRABAJADORES, BANCOS WHERE TRABAJADORES.BANCOCTS=BANCOS.CODBANCO AND TRABAJADORES.CODTRAB='" & RS!CODTRAB & "'", DBSYSTEM, adOpenKeyset, adLockOptimistic
            If RSAUX.RecordCount > 0 Then
                RS!BANCO = UCase(RSAUX!NOMBRE)
                RS!NROCUENTA = UCase(RSAUX!CTACTS)
                If RSAUX!MONE = "02" Then
                    RS!Moneda = "DOLARES"
                Else
                    RS!Moneda = "SOLES"
                End If
                RS!NOMBRES = UCase(RSAUX!ApePat) & " " & UCase(RSAUX!ApeMat) & " " & UCase(RSAUX!Empleado)
                RS!FECHAING = CDate(RSAUX!FECHAING)
                RS!DOCIDEN = RSAUX!DOCIDEN
                RS.Update
            End If
            TOTAL = TOTAL + RS!Importe
            RS.MoveNext
        Wend
        DBSYSTEM.Execute "UPDATE ##TMPCTS4 SET TOTAL=" & Round(DevuelveValor("SELECT SUM(IMPORTECTS) FROM  [##TMPCTS3" & VGL_COMPUTER & "]  WHERE CODTRAB='" & Codigo & "'", DBSYSTEM), 2) & " WHERE CODTRAB='" & Codigo & "'"
        TOTAL = 0
    End If
    If ExisteTablaAux(" [##TMP0001" & VGL_COMPUTER & "] ") Then DBSYSTEM.Execute "DROP TABLE  [##TMP0001" & VGL_COMPUTER & "] "
    DBSYSTEM.Execute "SELECT DISTINCT CODTRAB, NOMBRES, NROCUENTA, TOTAL, BANCO, FECHAING,MONEDA INTO  [##TMP0001" & VGL_COMPUTER & "]  FROM ##TMPCTS4 "
    DBSYSTEM.Execute "ALTER TABLE  [##TMP0001" & VGL_COMPUTER & "]  ADD LETRAS varchar(200)"
    DBSYSTEM.Execute "ALTER TABLE  [##TMP0001" & VGL_COMPUTER & "]  ADD REMUNERACION  Numeric(20,2) "
    Dim LETRAS As String
    Set RSAUX = New ADODB.Recordset
    RSAUX.Open " [##TMP0001" & VGL_COMPUTER & "] ", DBSYSTEM, adOpenKeyset, adLockOptimistic
    If RSAUX.RecordCount > 0 Then
        While Not RSAUX.EOF
            DBSYSTEM.Execute "UPDATE  [##TMP0001" & VGL_COMPUTER & "]  SET REMUNERACION=" & DevuelveValor("SELECT SUM(IMPORTE) FROM ##TMPCTS4 WHERE CODTRAB='" & RSAUX!CODTRAB & "'", DBSYSTEM) & " WHERE CODTRAB='" & RSAUX!CODTRAB & "'"
            LETRAS = NUMLET(Round(DevuelveValor("SELECT TOTAL FROM  [##TMP0001" & VGL_COMPUTER & "]  WHERE CODTRAB='" & RSAUX!CODTRAB & "'", DBSYSTEM), 2))
            DBSYSTEM.Execute "UPDATE  [##TMP0001" & VGL_COMPUTER & "]  SET LETRAS='*** " & LETRAS & " ***' WHERE CODTRAB='" & RSAUX!CODTRAB & "'"
            RSAUX.MoveNext
        Wend
    End If
    Dim GLOSA As String
    DBSYSTEM.Execute "DROP TABLE  [##TMPCTS3" & VGL_COMPUTER & "] "
    DBSYSTEM.Execute "DROP TABLE ##TMPCTS4"
    Dim RSEMPRESA As New ADODB.Recordset
    Dim FECHAS As String
    RSEMPRESA.Open "EMPRESA", DBSYSTEM, adOpenKeyset, adLockOptimistic
    If RSEMPRESA.RecordCount = 0 Then Exit Sub
    If ExisteTablaAux(" [##TEMPAUX" & VGL_COMPUTER & "] ") Then DBSYSTEM.Execute "DROP TABLE  [##TEMPAUX" & VGL_COMPUTER & "]  "
    DBSYSTEM.Execute "SELECT A.*,B.MESES,B.DIAS,C.CONCEPTO,C.IMPORTE INTO  [##TEMPAUX" & VGL_COMPUTER & "]  FROM  [##TMP0001" & VGL_COMPUTER & "]  A,PLANCTS B,DETALLECTS C  WHERE  B.CODIGO=C.CODIGO and B.CODTRAB = C.CODTRAB And A.CODTRAB = B.CODTRAB And B.Codigo = " & VPTRASPRM
    ':Fernando :12/02/2003
    With Reporte
        .Reset
        .ReportFileName = REGSISTEMA.REPORTES & "\pl_certificts.rpt"
        .Connect = "DSN=" & VGL_SERVERREP & ";User=sa;PWD=;DSQ=planilla_pp"
        .StoredProcParam(0) = " [##TEMPAUX" & VGL_COMPUTER & "] "
        .Destination = crptToWindow
        .WindowShowPrintBtn = True
        .WindowShowRefreshBtn = True
        .WindowShowSearchBtn = True
        .WindowShowPrintSetupBtn = True
        .WindowTitle = "CERTIFICADO - LIQUIDACION DE COMPESACION POR TIEMPO DE SERVICIOS - CTS."
        .Formulas(0) = "XEMPRESA='" & REGSISTEMA.EMPRESA & "'"
        .Formulas(2) = "XRUC='" & REGSISTEMA.RUC & "'"
        .Formulas(5) = "XDIRECCION='" & IIf(IsNull(RSEMPRESA!DIRECCIÓN), "  ", IIf(IsNull(RSEMPRESA!DISTRITO), "  ", UCase(RSEMPRESA!DIRECCIÓN) & " " & UCase(RSEMPRESA!DISTRITO))) & "'"
         GLOSA = IIf(IsNull(RSEMPRESA!DISTRITO), "  ", RSEMPRESA!DISTRITO) & "  " & Format(Date, "DDDD, DD MMMM YYYY")
        .Formulas(6) = "XGLOSA='" & GLOSA & "'"
         FECHAS = FECIN & " AL " & FECFI
        .Formulas(7) = "XINICIO='" & FECHAS & "'"
        .Formulas(8) = "XMES='" & UCase(Format(FECIN, "MMMM")) & "'"
        .Formulas(9) = "XTC='" & MDIPrincipal.BarraEstado.Panels(3).Text & "'"
        .Formulas(10) = "FechaDepo=DateValue (" & Year(DTPDepo.Value) & "," & Month(DTPDepo.Value) & "," & Day(DTPDepo.Value) & ")"
        .Formulas(11) = "FechaEmi=DateValue (" & Year(DTPEmi.Value) & "," & Month(DTPEmi.Value) & "," & Day(DTPEmi.Value) & ")"
        If .Status <> 2 Then .Action = 1
    End With
    Screen.MousePointer = 1
End Sub
Private Sub Command2_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    If ExisteTablaAux(" [##TMPCTS3" & VGL_COMPUTER & "] ") Then DBSYSTEM.Execute "DROP TABLE  [##TMPCTS3" & VGL_COMPUTER & "] "
        DBSYSTEM.Execute "CREATE TABLE  [##TMPCTS3" & VGL_COMPUTER & "]  (CODTRAB VARCHAR(8), NOMBRES VARCHAR(100), IMPORTECTS  Numeric(20,2) , MESES BIT, DIAS BIT, FECHAING DATETIME)"
    If ExisteTablaAux("##TMPCTS4") Then DBSYSTEM.Execute "DROP TABLE ##TMPCTS4"
        DBSYSTEM.Execute "CREATE TABLE ##TMPCTS4 (CODTRAB VARCHAR(8), CONCEPTO VARCHAR(35), IMPORTE  Numeric(20,2) , INDTIPO BIT, BANCO VARCHAR(50), NROCUENTA VARCHAR(50), NOMBRES VARCHAR(100), FECHAING DATETIME, DOCIDEN VARCHAR(15))"
        DBSYSTEM.Execute "CREATE INDEX CODTRAB ON ##TMPCTS4 (CODTRAB) "
    xPeriodo = DevuelveValor("SELECT NOMBRE FROM CTS WHERE CODIGO=" & VPTRASPRM, DBSYSTEM)
    FECIN = DevuelveValor("SELECT FECHAINI FROM CTS WHERE CODIGO=" & VPTRASPRM, DBSYSTEM)
    FECFI = DevuelveValor("SELECT FECHAFIN FROM CTS WHERE CODIGO=" & VPTRASPRM, DBSYSTEM)
    DBSYSTEM.Execute "INSERT INTO  [##TMPCTS3" & VGL_COMPUTER & "]  SELECT CODTRAB, LTRIM(RTRIM(NOMBRES)), IMPORTECTS, MESES, DIAS, FECHAING FROM PLANCTS WHERE CODIGO=" & VPTRASPRM
    DBSYSTEM.Execute "INSERT INTO ##TMPCTS4 (CODTRAB, CONCEPTO, IMPORTE, INDTIPO) SELECT CODTRAB, CONCEPTO, IMPORTE, INDTIPO FROM DETALLECTS WHERE CODIGO=" & VPTRASPRM
End Sub

