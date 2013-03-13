VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form FrmResuDeb 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Resumen de Debitos Por Meses"
   ClientHeight    =   2355
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4455
   Icon            =   "FrmResuDeb.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2355
   ScaleWidth      =   4455
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Opciones"
      Height          =   825
      Left            =   90
      TabIndex        =   8
      Top             =   0
      Width           =   4200
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
         Left            =   1575
         TabIndex        =   10
         Top             =   375
         Width           =   1050
      End
      Begin VB.OptionButton xTodos 
         Caption         =   "Todos"
         Height          =   300
         Left            =   2910
         TabIndex        =   9
         Top             =   315
         Width           =   1050
      End
   End
   Begin VB.Frame FraFecha 
      Caption         =   "Pör Mes"
      Height          =   1290
      Left            =   75
      TabIndex        =   3
      Top             =   930
      Width           =   2835
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
      Left            =   3105
      TabIndex        =   2
      Top             =   1185
      Width           =   1140
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   330
      Left            =   3120
      TabIndex        =   1
      Top             =   1605
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
Attribute VB_Name = "FrmResuDeb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Function VALIDAR() As Boolean
VALIDAR = True
    If xFechaIni > xFechaFin Then
        MsgBox "LA FECHA DESDE NO PUEDE SER MAYOR QUE LA FECHA HASTA ", vbExclamation
        Exit Function
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
    Dim CAMPO As String
    Dim FECHAS() As String
    Dim TIPO As String
    If VALIDAR Then Exit Sub 'VALIDANDO EL INGRESO
    OPCION = ""
    Screen.MousePointer = 11
    INTO = " INTO  [##TMPCTAGRUP" & VGL_COMPUTER & "]  "
    FECHAINICIO = ""
    FECHAFIN = ""
    FECHAINICIO = Trim("#" & Month(xFechaIni) & "/01/" & Year(xFechaIni) & "#")
    FECHAFIN = Trim("#" & Month(xFechaFin) & "/01/" & Year(xFechaFin) & "#")
    OPCFECHAS = " AND NOMBOL.MES BETWEEN " & FECHAINICIO & " AND " & FECHAFIN
    PERIODOINI = Format(DateValue("01/" & Month(xFechaIni) & "/" & Year(xFechaIni)), "MMMM  -  YYYY")
    PERIODOFIN = Format(DateValue("01/" & Month(xFechaFin) & "/" & Year(xFechaFin)), "MMMM  -  YYYY")
    RESFECHAS = DateDiff("M", xFechaIni, xFechaFin)
    ReDim FECHAS(RESFECHAS)
    If RESFECHAS > 9 Then
        MsgBox ("EL INTERVALO DE LA FECHA DESDE CON LA FECHA HASTA DEBE SER < = 10")
        Screen.MousePointer = 1
        Exit Sub
    End If
    RESFECHAS = RESFECHAS + 1
    Dim OPC1 As String
    OPC1 = ""
    TIPO = ""
    If XIngresos.Value Then
        OPCION = " WHERE TIPO=1"
        OPC1 = " AND PAGOSCTA.TIPO=1"
        TIPO = "INGRESOS"
    End If
    If xEgresos.Value Then
        OPCION = " WHERE TIPO=2"
        OPC1 = " AND PAGOSCTA.TIPO=2"
        TIPO = "EGRESOS"
    End If
    If xTodos.Value Then
        OPCION = ""
        OPC1 = ""
        TIPO = "TODOS"
    End If
    Dim CERO As String
    CERO = Format(0, "0.00")
    Dim RUTA As String
    'ARMANDO LA TABLA RESUMEN CON 12 MESES
    SqlCad.Text = " " & _
    "   SELECT CODGRUPO, NOMBRE,X1=CAST(0 AS FLOAT),X2=CAST(0 AS FLOAT),X3=CAST(0 AS FLOAT),X4=CAST(0 AS FLOAT), " & _
    " X5=CAST(0 AS FLOAT),X6=CAST(0 AS FLOAT),X7=CAST(0 AS FLOAT),X8=CAST(0 AS FLOAT), " & _
    " X9=CAST(0 AS FLOAT),X10=CAST(0 AS FLOAT),X11=CAST(0 AS FLOAT),X12=CAST(0 AS FLOAT) " & _
    INTO & "  FROM CTAGRUPO  " & OPCION
    
    If ExisteTablaAux(" [##TMPCTAGRUP" & VGL_COMPUTER & "] ") Then DBSYSTEM.Execute "DROP TABLE  [##TMPCTAGRUP" & VGL_COMPUTER & "] "
    DBSYSTEM.Execute SqlCad.Text
    
    Dim XAUXFECHA As Date
    Dim RSCTAGRUP As New ADODB.Recordset
    Set RSCTAGRUP = New ADODB.Recordset
       
    RSCTAGRUP.Open "SELECT * FROM  [##TMPCTAGRUP" & VGL_COMPUTER & "]  ORDER BY CODGRUPO", DBAUXCOM, adOpenDynamic, adLockOptimistic
    
    XAUXFECHA = xFechaIni
    Dim RSAUX As New ADODB.Recordset
    
    For I = 1 To RESFECHAS
        Set RSAUX = New ADODB.Recordset
        xMes = Month(XAUXFECHA): xAnno = Year(XAUXFECHA)
        FECHAS(I - 1) = Format(DateValue("01/" & Trim(Str(xMes)) & "/" & Trim(Str(xAnno))), "MMM-YYYY")
        SqlCad.Text = ""
        SqlCad.Text = " " & _
        "SELECT MOVICTA.CODGRUPO, SUM(PAGOSCTA.MONTO) AS SMONTO " & _
        " FROM MOVICTA, NOMBOL, PAGOSCTA " & _
        " WHERE (((PAGOSCTA.CODNOMBOL)=[NOMBOL].[CODIGO]) " & _
        " AND ((PAGOSCTA.CODMOV)=[MOVICTA].[CODMOV])) AND " & _
        " MONTH(NOMBOL.MES)= " & Str(xMes) & " AND YEAR(NOMBOL.MES)= " & Str(xAnno) & _
          OPC1 & " GROUP BY MOVICTA.CODGRUPO ORDER BY MOVICTA.CODGRUPO "
        RSAUX.Open SqlCad.Text, DBSYSTEM, adOpenKeyset, adLockOptimistic
        
        If RSAUX.RecordCount <> 0 Then
            RSAUX.MoveFirst
            For j = 1 To RSAUX.RecordCount
                RSCTAGRUP.MoveFirst
                Do While Not (RSCTAGRUP!CODGRUPO = RSAUX!CODGRUPO)
                    If Not RSAUX.EOF Then
                        RSCTAGRUP.MoveNext
                      Else: Exit Do
                    End If
                Loop
                CAMPO = "X" & Trim(Str(I))
                RSCTAGRUP.Fields(CAMPO).Value = RSAUX!SMONTO
                RSCTAGRUP.Update
                RSAUX.MoveNext
            Next
         End If
        XAUXFECHA = DateAdd("M", I, xFechaIni)
   Next

    DBSTARPLAN.Execute "EXECUTE [ASISTMP] ' [##TMPCTAGRUP" & VGL_COMPUTER & "] '"
   
    With Reporte
        .ReportFileName = REGSISTEMA.REPORTES & "PLAN0018.RPT"
        .WindowTitle = "PLAN0018 - CONSOLIDADO POR MES DE LOS DEBITOS DE CTA. CTE."
        .Connect = "DSN=" & VGL_SERVERREP & ";User=sa;PWD=;DSQ=planilla_pp"
        .StoredProcParam(0) = " [##TMPCTAGRUP" & VGL_COMPUTER & "] "
        .Destination = crptToWindow
        .WindowShowPrintBtn = True
        .WindowShowRefreshBtn = True
        .WindowShowSearchBtn = True
        .WindowShowPrintSetupBtn = True
        .Formulas(0) = "XEMPRESA='" & REGSISTEMA.EMPRESA & "'"
        .Formulas(1) = "XRUC='" & REGSISTEMA.RUC & "'"
        .Formulas(2) = "XFECHINI='" & PERIODOINI & "'"
        .Formulas(3) = "XFECHFIN='" & PERIODOFIN & "'"
        .Formulas(4) = "XTIPO='" & TIPO & "'"
        .Formulas(5) = "XHORA='" & Format(Time, "HH:MM") & "'"
        For I = 0 To UBound(FECHAS)
            .Formulas(6 + I) = "X" & Trim(Str(I + 1)) & "='" & FECHAS(I) & "'"
        Next
        If .Status <> 2 Then .Action = 1
    End With
    Screen.MousePointer = 1
End Sub


Private Sub CMDCANCELAR_CLICK()
    Unload Me
End Sub


