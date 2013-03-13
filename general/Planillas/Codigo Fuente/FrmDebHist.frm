VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form FrmDebHist 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Historial Debitos x Trabajador"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4275
   Icon            =   "FrmDebHist.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   4275
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Moneda"
      Height          =   720
      Left            =   90
      TabIndex        =   12
      Top             =   750
      Width           =   4080
      Begin VB.OptionButton OpME 
         Caption         =   "Moneda E&xtranjera"
         Height          =   255
         Left            =   2280
         TabIndex        =   14
         Top             =   315
         Width           =   1695
      End
      Begin VB.OptionButton OpMN 
         Caption         =   "Moneda &Nacional"
         Height          =   240
         Left            =   120
         TabIndex        =   13
         Top             =   300
         Width           =   1740
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Opciones"
      Height          =   720
      Left            =   90
      TabIndex        =   8
      Top             =   0
      Width           =   4095
      Begin VB.OptionButton XIngresos 
         Caption         =   "&Ingresos"
         Height          =   240
         Left            =   120
         TabIndex        =   11
         Top             =   345
         Width           =   1050
      End
      Begin VB.OptionButton xEgresos 
         Caption         =   "&Egresos"
         Height          =   210
         Left            =   1545
         TabIndex        =   10
         Top             =   315
         Width           =   1050
      End
      Begin VB.OptionButton xTodos 
         Caption         =   "Am&bos"
         Height          =   300
         Left            =   2910
         TabIndex        =   9
         Top             =   300
         Width           =   1050
      End
   End
   Begin VB.Frame FraFecha 
      Height          =   1230
      Left            =   75
      TabIndex        =   3
      Top             =   1560
      Width           =   2835
      Begin VB.CheckBox ChkFech 
         Alignment       =   1  'Right Justify
         Caption         =   "Por Fechas"
         Height          =   255
         Left            =   75
         TabIndex        =   15
         Top             =   0
         Width           =   1305
      End
      Begin MSComCtl2.DTPicker xFechaFin 
         Height          =   300
         Left            =   915
         TabIndex        =   4
         Top             =   750
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   393216
         Format          =   62390273
         CurrentDate     =   36691
      End
      Begin MSComCtl2.DTPicker xFechaIni 
         Height          =   300
         Left            =   915
         TabIndex        =   5
         Top             =   375
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   393216
         Format          =   62390273
         CurrentDate     =   36691
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         Height          =   195
         Left            =   150
         TabIndex        =   7
         Top             =   855
         Width           =   420
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         Height          =   195
         Left            =   165
         TabIndex        =   6
         Top             =   390
         Width           =   465
      End
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   345
      Left            =   3030
      TabIndex        =   2
      Top             =   1830
      Width           =   1140
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   330
      Left            =   3045
      TabIndex        =   1
      Top             =   2310
      Width           =   1125
   End
   Begin VB.TextBox SqlCad 
      Height          =   285
      Left            =   690
      TabIndex        =   0
      Text            =   "SqlCad"
      Top             =   1620
      Visible         =   0   'False
      Width           =   1185
   End
   Begin Crystal.CrystalReport Reporte 
      Left            =   90
      Top             =   1545
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "FrmDebHist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Function VALIDAR() As Boolean
VALIDAR = True
  If ChkFech.Value = 1 Then
    If xFechaIni > xFechaFin Then
        MsgBox "LA FECHA DESDE NO PUEDE SER MAYOR QUE LA FECHA HASTA ", vbExclamation
        Exit Function
    End If
  End If
VALIDAR = False
End Function
Private Sub CMDACEPTAR_CLICK()
    Dim INTO As String
    Dim FECHAINICIO As String
    Dim FECHAFIN As String
    Dim OPCFECHAS As String, PERIODOINI As String, PERIODOFIN As String
    Dim X As Integer, RESFECHAS As Long, I As Integer, j As Integer
    Dim xMes As Integer, xAnno As Integer
    Dim TIPO As String
    If VALIDAR Then Exit Sub 'VALIDANDO EL INGRESO
    
    Screen.MousePointer = 11
    'INTO = " INTO [" & APP.PATH & "\BDAUXCOM.MDB].TMPDEBHIST  "
    INTO = " INTO  [##TMPDEBHIST" & VGL_COMPUTER & "]  "
    
    FECHAINICIO = ""
    FECHAFIN = ""
    PERIODOINI = ""
    PERIODOFIN = ""
    
    If ChkFech.Value = 1 Then
        FECHAINICIO = Trim("'" & Month(xFechaIni) & "/01/" & Year(xFechaIni) & "'")
        FECHAFIN = Trim("'" & Month(xFechaFin) & "/01/" & Year(xFechaFin) & "'")
        OPCFECHAS = " AND NOMBOL.MES BETWEEN " & FECHAINICIO & " AND " & FECHAFIN
        PERIODOINI = Format(DateValue("01/" & Month(xFechaIni) & "/" & Year(xFechaIni)), "MMMM  -  YYYY")
        PERIODOFIN = Format(DateValue("01/" & Month(xFechaFin) & "/" & Year(xFechaFin)), "MMMM  -  YYYY")
    End If
    
    Dim OPC1 As String
    OPC1 = "":    TIPO = ""
    If XIngresos.Value Then
        OPC1 = " AND PAGOSCTA.TIPO=1"
        TIPO = "INGRESOS"
    End If
    If xEgresos.Value Then
        OPC1 = " AND PAGOSCTA.TIPO=2"
        TIPO = "EGRESOS"
    End If
    If xTodos.Value Then
        OPC1 = ""
        TIPO = "TODOS"
    End If
    
    Dim OPCMONEDA As String
    Dim TIPOMON As String
    
    OPCMONEDA = "": TIPOMON = ""
    
    If OpMN.Value Then
        OPCMONEDA = " AND MOVICTA.MONEDA=0"
        TIPOMON = "MONEDA NACIONAL(S/.)"
    End If
    If OpME.Value Then
        OPCMONEDA = " AND MOVICTA.MONEDA=1"
        TIPOMON = "MONEDA EXTRANJERA(US$)"
    End If
    Dim RUTATRAB As String
    RUTATRAB = " IN '" & REGSISTEMA.PATHEMPRESA & "\PLANILLA.MDB' "
    
    'ARMANDO LA CONSULTA
    
    SqlCad.Text = " " & _
    "SELECT  NOMBOL.NOMBRE, MOVICTA.CODMOV,MOVICTA.FECHAINI, MOVICTA.DESCRIPCION,CTAGRUPO.CODGRUPO, " & _
    "CTAGRUPO.NOMBRE AS CTA, TRABAJADORES.CODTRAB, " & _
    "LTRIM([TRABAJADORES].[APEPAT])+' '+LTRIM([TRABAJADORES].[APEMAT])+' '+LTRIM([TRABAJADORES].[NOMBRE]) AS NOMBRES, " & _
    "PAGOSCTA.TIPOBOLETA, " & _
    "PAGO = CASE [PAGOSCTA].[TIPOBOLETA] WHEN 'B' THEN 'PAGO EN BOLETA REMUNERACIONES  '+LTRIM([NOMBOL].[NOMBRE])+' '+LTRIM(STR(YEAR([NOMBOL].[MES]))) ELSE 'PAGO EN ADELANTO  '+LTRIM([NOMBOL].[NOMBRE])+' '+LTRIM(STR(YEAR([NOMBOL].[MES]))) END, PAGOSCTA.MONTO, MOVICTA.CAPITAL " & _
    INTO & _
    "FROM CTAGRUPO, MOVICTA, NOMBOL, PAGOSCTA, TRABAJADORES " & _
    "WHERE (((PAGOSCTA.CODNOMBOL) = [NOMBOL].[CODIGO]) AND " & _
    "((PAGOSCTA.CODMOV) = [MOVICTA].[CODMOV]) AND ((MOVICTA.CODTRAB) " & _
    "= [TRABAJADORES].[CODTRAB]) AND ((MOVICTA.CODGRUPO) = [CTAGRUPO].[CODGRUPO])) " & OPC1 & OPCMONEDA & OPCFECHAS
    
    If ExisteTablaAux(" [##TMPDEBHIST" & VGL_COMPUTER & "] ") Then DBSYSTEM.Execute "DROP TABLE  [##TMPDEBHIST" & VGL_COMPUTER & "] "
    DBSYSTEM.Execute SqlCad.Text, X
    If X = 0 Then
        MsgBox "NO SE ENCONTRARÓN REGISTROS", vbExclamation
        Screen.MousePointer = 1
        Exit Sub
    End If
    With Reporte
        .ReportFileName = REGSISTEMA.REPORTES & "PLAN0019.RPT"
        .WindowTitle = "PLAN0019 - HISTORIAL - DEBITOS CTA. CTE."
        .Connect = "DSN=" & VGL_SERVERREP & ";User=sa;PWD=;DSQ=planilla_pp"
        .StoredProcParam(0) = " [##TMPDEBHIST" & VGL_COMPUTER & "] "
        .Destination = crptToWindow
        .WindowShowPrintBtn = True
        .WindowShowRefreshBtn = True
        .WindowShowSearchBtn = True
        .WindowShowPrintSetupBtn = True
        .Formulas(0) = "XEMPRESA='" & REGSISTEMA.EMPRESA & "'"
        .Formulas(1) = "XRUC='" & REGSISTEMA.RUC & "'"
        If PERIODOINI <> "" And PERIODOFIN <> "" Then
            .Formulas(2) = "XFECHINI='" & Format(DateValue("01/" & PERIODOINI), "MMMM - YYYY") & "'"
            .Formulas(3) = "XFECHFIN='" & Format(DateValue("01/" & PERIODOFIN), "MMMM - YYYY") & "'"
          Else:
            .Formulas(2) = "XFECHINI='" & PERIODOINI & "'"
            .Formulas(3) = "XFECHFIN='" & PERIODOFIN & "'"
         End If
        .Formulas(4) = "XTIPO='" & TIPO & "'"
        .Formulas(5) = "XHORA='" & Format(Time, "HH:MM") & "'"
        .Formulas(6) = "XMONEDA='" & TIPOMON & "'"
        If .Status <> 2 Then .Action = 1
    End With
    Screen.MousePointer = 1
End Sub

Private Sub CMDCANCELAR_CLICK()
    Unload Me
End Sub

Private Sub Form_Load()
    XIngresos.Value = True
    OpMN.Value = True
End Sub


