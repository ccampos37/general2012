VERSION 5.00
Object = "{C5FF36B7-A67E-11D3-9D9D-E6F193F7F854}#9.0#0"; "ApliCTxt.ocx"
Begin VB.Form FrPagReg 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reporte - Planilla de pago de aportes"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5760
   Icon            =   "FrPagReg.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   5760
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   840
      Left            =   60
      TabIndex        =   10
      Top             =   2235
      Width           =   5610
      Begin AplisetControlText.Aplitext x1Pag 
         Height          =   315
         Left            =   1440
         TabIndex        =   2
         Top             =   285
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         Text            =   "0"
         Entero          =   -1  'True
         SinBlancos      =   -1  'True
         DigitRound      =   0
         TipoDato        =   "N"
      End
      Begin AplisetControlText.Aplitext xdPag 
         Height          =   315
         Left            =   4200
         TabIndex        =   3
         Top             =   285
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         Text            =   "0"
         Entero          =   -1  'True
         SinBlancos      =   -1  'True
         DigitRound      =   0
         TipoDato        =   "N"
      End
      Begin VB.Label Label1 
         Caption         =   "Reg primera Pag."
         Height          =   315
         Left            =   60
         TabIndex        =   12
         Top             =   345
         Width           =   1410
      End
      Begin VB.Label Label2 
         Caption         =   "Reg demas Pag."
         Height          =   300
         Left            =   2865
         TabIndex        =   11
         Top             =   345
         Width           =   1290
      End
   End
   Begin VB.ComboBox CmbColum 
      Height          =   315
      ItemData        =   "FrPagReg.frx":0442
      Left            =   3315
      List            =   "FrPagReg.frx":044C
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1845
      Visible         =   0   'False
      Width           =   2385
   End
   Begin VB.CheckBox ChkOrder 
      Caption         =   "Ordenar al Imprimir"
      Height          =   285
      Left            =   105
      TabIndex        =   0
      Top             =   1875
      Width           =   1875
   End
   Begin VB.Frame Frame1 
      Caption         =   "Detalle"
      Height          =   1605
      Left            =   60
      TabIndex        =   7
      Top             =   105
      Width           =   5640
      Begin VB.Label Label5 
         Caption         =   "Nota.- Si usted ordena, no se considerará los numeros impares para la cantidad de registros en la primera pag."
         Height          =   390
         Left            =   1080
         TabIndex        =   13
         Top             =   1005
         Width           =   4125
      End
      Begin VB.Label Label3 
         Caption         =   $"FrPagReg.frx":0476
         Height          =   615
         Left            =   1065
         TabIndex        =   8
         Top             =   315
         Width           =   4200
         WordWrap        =   -1  'True
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   135
         Picture         =   "FrPagReg.frx":04FE
         Stretch         =   -1  'True
         Top             =   285
         Width           =   480
      End
   End
   Begin VB.CommandButton Command3 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   3720
      TabIndex        =   6
      Top             =   3195
      Width           =   1500
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Ver &Demas Pag."
      Enabled         =   0   'False
      Height          =   375
      Left            =   2145
      TabIndex        =   5
      Top             =   3195
      Width           =   1500
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ver &Primera Pag."
      Height          =   375
      Left            =   555
      TabIndex        =   4
      Top             =   3195
      Width           =   1500
   End
   Begin VB.Label LbColum 
      Caption         =   "Columna"
      Height          =   255
      Left            =   2535
      TabIndex        =   9
      Top             =   1920
      Visible         =   0   'False
      Width           =   825
   End
End
Attribute VB_Name = "FrPagReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim FLAG As Boolean

Private Sub CHKORDER_Click()
    If ChkOrder.Value = 1 Then
        LbColum.Visible = True
        CmbColum.Visible = True
        CmbColum.ListIndex = 0
        Else:
        LbColum.Visible = False
        CmbColum.Visible = False
    End If
End Sub

Private Sub CHKORDER_KeyDown(KEYCODE As Integer, Shift As Integer)
If KEYCODE Then SendKeys "{TAB}"
End Sub

Private Sub Command1_Click()
    Dim TOP As Integer
    TOP = Val(x1Pag.Text)
    
    If TOP = 0 Then
        MsgBox "Por lo menos debe imprimir un registro", vbExclamation
        x1Pag.SetFocus
        Exit Sub
    End If
    If TOP >= 19 Then
        MsgBox "Pueden ser hasta 19 registros en la primera hoja", vbInformation
        Exit Sub
    End If
    
    
    Dim ORDENAR As String
    
    Screen.MousePointer = 11
    ORDENAR = ""
    If ChkOrder.Value = 1 Then
        If CmbColum.ListIndex = 0 Then ORDENAR = "ORDER BY CODTRAB"
        If CmbColum.ListIndex = 1 Then ORDENAR = "ORDER BY APEPAT,APEMAT,NOMBRES"
      Else:
        ORDENAR = ""
    End If
    
    
    With frPlanAFP
        SaveSetting App.CompanyName, "AFP", "RESPON", .xResponsable.Text
        SaveSetting App.CompanyName, "AFP", "DEPARTAM", .xDepartamento.Text
        SaveSetting App.CompanyName, "AFP", "CTACTE", .xCtaBanco.Text
        SaveSetting App.CompanyName, "AFP", "BANCO", .xBanco.Text
        SaveSetting App.CompanyName, "AFP", "TIPOCTA", .xTipoCta.Text
    'INICIO DEL REPORTE
    End With
    
    Dim RSAUX As New ADODB.Recordset
    Dim SREMUASEG As Double, SAPOROBLI As Double
    Dim SAPORVOLT As Double, SAPORVOLE As Double
    Dim SAPOREMP As Double
    Dim XSUMAFP As Double, XTOTALFONDO As Double
    Dim S1M1 As Double, STOT1 As Double
    Dim S1M2 As Double, STOT2 As Double
    Dim XSUMRC As Double, STOTALRC As Double
    Dim SSEGUROS As Double, SCOMISION As Double
    
    If ExisteTablaAux(" [##_TMPPLANAFP2" & VGL_COMPUTER & "] ") Then
        If ExisteTablaAux(" [##_TMPPLANAFP" & VGL_COMPUTER & "] ") Then DBSYSTEM.Execute "DROP TABLE  [##_TMPPLANAFP" & VGL_COMPUTER & "] "
        DBSYSTEM.Execute "SELECT * INTO  [##_TMPPLANAFP" & VGL_COMPUTER & "]  FROM  [##_TMPPLANAFP2" & VGL_COMPUTER & "] "
    End If
    
    RSAUX.Open "SELECT * FROM  [##_TMPPLANAFP" & VGL_COMPUTER & "]  WHERE CODAFP ='" & _
                frPlanAFP.RSRESAFP!Codigo & "'", DBSYSTEM, adOpenKeyset, adLockOptimistic
    SREMUASEG = SAPOROBLI = 0
    SAPORVOLT = SAPORVOLE = 0
    SAPOREMP = XSUMAFP = XTOTALFONDO = S1M1 = STOT1 = S1M2 = 0
    STOT2 = XSUMRC = STOTALRC = SSEGUROS = SCOMISION = 0
    
    RSAUX.MoveFirst
    Do While Not RSAUX.EOF
       SREMUASEG = SREMUASEG + IIf(IsNull(RSAUX!REMUASEG), 0, RSAUX!REMUASEG)
       SAPOROBLI = SAPOROBLI + IIf(IsNull(RSAUX!APOROBLI), 0, RSAUX!APOROBLI)
       SAPORVOLT = SAPORVOLT + IIf(IsNull(RSAUX!APORVOLT), 0, RSAUX!APORVOLT)
       SAPORVOLE = SAPORVOLE + IIf(IsNull(RSAUX!APORVOLE), 0, RSAUX!APORVOLE)
       SAPOREMP = SAPOREMP + IIf(IsNull(RSAUX!APOREMP), 0, RSAUX!APOREMP)
       XSUMAFP = IIf(IsNull(RSAUX!APOROBLI), 0, RSAUX!APOROBLI) + IIf(IsNull(RSAUX!APORVOLT), 0, RSAUX!APORVOLT) + IIf(IsNull(RSAUX!APORVOLE), 0, RSAUX!APORVOLE) + IIf(IsNull(RSAUX!APOREMP), 0, RSAUX!APOREMP)
       XTOTALFONDO = XTOTALFONDO + XSUMAFP
       SSEGUROS = SSEGUROS + IIf(IsNull(RSAUX!SEGUROS), 0, RSAUX!SEGUROS): SCOMISION = SCOMISION + IIf(IsNull(RSAUX!COMISION), 0, RSAUX!COMISION)
       XSUMRC = IIf(IsNull(RSAUX!SEGUROS), 0, RSAUX!SEGUROS) + IIf(IsNull(RSAUX!COMISION), 0, RSAUX!COMISION)
       STOTALRC = STOTALRC + XSUMRC
       RSAUX.MoveNext
    Loop
    S1M1 = XTOTALFONDO * Val(frPlanAFP.xInteres.Text)
    S1M2 = STOTALRC * Val(frPlanAFP.xInteres.Text)
    STOT1 = S1M1 + XTOTALFONDO
    STOT2 = S1M2 + STOTALRC
    
    
    Set RSAUX = Nothing
    If ExisteTablaAux(" [##_TMPPLANAFP2" & VGL_COMPUTER & "] ") Then DBSYSTEM.Execute "DROP TABLE  [##_TMPPLANAFP2" & VGL_COMPUTER & "] "
    DBSYSTEM.Execute "SELECT * INTO  [##_TMPPLANAFP2" & VGL_COMPUTER & "]  FROM  [##_TMPPLANAFP" & VGL_COMPUTER & "]  "
    
    If ExisteTablaAux(" [##_TMPPLANAFP1" & VGL_COMPUTER & "] ") Then DBSYSTEM.Execute "DROP TABLE  [##_TMPPLANAFP1" & VGL_COMPUTER & "] "
    DBSYSTEM.Execute "SELECT * INTO  [##_TMPPLANAFP1" & VGL_COMPUTER & "]  FROM  [##_TMPPLANAFP" & VGL_COMPUTER & "]  " & _
                     "WHERE CODAFP ='" & frPlanAFP.RSRESAFP!Codigo & "'"
    
    If ExisteTablaAux(" [##_TMPPLANAFP" & VGL_COMPUTER & "] ") Then DBSYSTEM.Execute "DROP TABLE  [##_TMPPLANAFP" & VGL_COMPUTER & "] "
    DBSYSTEM.Execute "SELECT TOP " & Str(TOP) & " * INTO  [##_TMPPLANAFP" & VGL_COMPUTER & "]  FROM  [##_TMPPLANAFP1" & VGL_COMPUTER & "]  " & _
                     ORDENAR
                     
    If ExisteTablaAux(" [##_TMPPLANAFPREST" & VGL_COMPUTER & "] ") Then DBSYSTEM.Execute "DROP TABLE  [##_TMPPLANAFPREST" & VGL_COMPUTER & "] "
    DBSYSTEM.Execute "SELECT *,1 AS PAG INTO  [##_TMPPLANAFPREST" & VGL_COMPUTER & "]  FROM  [##_TMPPLANAFP1" & VGL_COMPUTER & "]    " & _
                     "WHERE INUMBOL NOT IN (SELECT INUMBOL FROM  [##_TMPPLANAFP" & VGL_COMPUTER & "]  ) " & _
                     ORDENAR
                
    Dim RsEmp As New ADODB.Recordset
    RsEmp.Open "EMPRESA", DBSYSTEM, adOpenStatic
    With frPlanAFP.RptAFP
        .Reset
        .ReportFileName = REGSISTEMA.REPORTES & "PLAN0031.RPT"
        .Connect = "DSN=" & VGL_SERVER & ";User=sa;PWD=;DSQ=" & REGSISTEMA.BASESQL & ""
        .Destination = crptToWindow
        .StoredProcParam(0) = "[##_TMPPLANAFP" & VGL_COMPUTER & "]"
        .WindowTitle = "PLAN0031 - PRIMERA HOJA - PLANILLA DE PAGO DE APORTES"
        .WindowShowPrintBtn = True
        .WindowShowRefreshBtn = True
        .WindowShowSearchBtn = True
        .WindowShowPrintSetupBtn = True
        .ReplaceSelectionFormula "{ASISTMP.CODAFP}='" & frPlanAFP.RSRESAFP!Codigo & "'"
        .Formulas(0) = "NOMAFP='" & frPlanAFP.RSRESAFP!NOMAFP & "'"
        .Formulas(1) = "NOMEMPRESA='" & REGSISTEMA.EMPRESA & "'"
        .Formulas(2) = "NUMPLANILLA='" & frPlanAFP.RSRESAFP!NUMPLANILLA & "'"
        .Formulas(3) = "APEPAT='" & Trim(RsEmp!RL_APEPAT) & " " & Trim(RsEmp!RL_APEMAT) & " " & RsEmp!RL_NOMBRE & "'"
        .Formulas(4) = "AREAURBANA='" & RsEmp!AREAURBANA & "'"
        .Formulas(5) = "BANCO='" & frPlanAFP.xBanco.Text & "'"
        .Formulas(6) = "CTABANCO='" & frPlanAFP.xCtaBanco.Text & "'"
        .Formulas(7) = "DEPARTAMENTO='" & RsEmp!DEPARTAMENTO & "'"
        .Formulas(8) = "DIRECCION='" & RsEmp!DIRECCIÓN & "'"
        .Formulas(9) = "DISTRITO='" & RsEmp!DISTRITO & "'"
        .Formulas(10) = "ELABORADOR='" & frPlanAFP.xResponsable.Text & "'"
        .Formulas(11) = "INTERIOR='" & RsEmp!INTERIOR & "'"
        .Formulas(12) = "MESPAGO='" & Right(frPlanAFP.dgAFPs.Caption, 7) & "'"
        .Formulas(13) = "PROVINCIA='" & RsEmp!PROVINCIA & "'"
        .Formulas(14) = "RUC='" & REGSISTEMA.RUC & "'"
        .Formulas(15) = "TELEFONO='" & RsEmp!TELEFONO1 & "'"
        .Formulas(16) = "TIPOCTA='" & frPlanAFP.xTipoCta.Text & "'"
        .Formulas(17) = "TIPODOC='" & RsEmp!RL_TIPODOC & "'"
        .Formulas(18) = "NUMDOC='" & RsEmp!RL_DOCUMENTO & "'"
        If frPlanAFP.f1Tipo(0).Value Then  'PAGO AL FONDO DE PENSIONES EN EFECTIVO
            .Formulas(19) = "F1TIPO='X'"
            .Formulas(20) = "F1TIPO2=''" 'CON CHEQUE
            .Formulas(21) = "F1CHEQUE=''"
            .Formulas(22) = "F1BANCO=''"
        Else
            .Formulas(19) = "F1TIPO2='X'" 'CON CHEQUE
            .Formulas(20) = "F1TIPO=''"
            .Formulas(21) = "F1CHEQUE='" & frPlanAFP.xf1Cheque.Text & "'"
            .Formulas(22) = "F1BANCO='" & frPlanAFP.xf1Banco.Text & "'"
        End If
        'NUMERO DE HOJAS ADICIONALES, FALTA CONFIGURACIÓN
        .Formulas(23) = "HOJASADIC=1"
        If frPlanAFP.f1Tipo(0).Value Then  'PAGO A LA AFP EN EFECTIVO
            .Formulas(24) = "F2EFEC='X'"
            .Formulas(25) = "F2CHEQUE=''"
            .Formulas(26) = "F2BANCO=''"
            .Formulas(27) = "F2TIPO2=''"
        Else
            .Formulas(24) = "F2TIPO2='X'" 'CON CHEQUE
            .Formulas(25) = "F2CHEQUE='" & frPlanAFP.xf2Cheque.Text & "'"
            .Formulas(26) = "F2BANCO='" & frPlanAFP.xf2Banco.Text & "'"
            .Formulas(27) = "F2EFEC=''"
        End If
        If frPlanAFP.fTipoPago(0).Value Then
            .Formulas(28) = "TP1='X'"
            .Formulas(29) = "TP2=''"
            .Formulas(30) = "TP3=''"
            .Formulas(31) = "TP4=''"
            .Formulas(32) = "TPLIQPLAN=''"
            .Formulas(33) = "TPREGPLAN=''"
            .Formulas(34) = "INTERES=0"
        End If
        If frPlanAFP.fTipoPago(1).Value Then
            .Formulas(28) = "TP1=''"
            .Formulas(29) = "TP2='X'"
            .Formulas(30) = "TP3=''"
            .Formulas(31) = "TP4=''"
            .Formulas(32) = "TPLIQPLAN=''"
            .Formulas(33) = "TPREGPLAN=''"
            .Formulas(34) = "INTERES=" & frPlanAFP.xInteres.Text
        End If
        If frPlanAFP.fTipoPago(2).Value Then
            .Formulas(28) = "TP1=''"
            .Formulas(29) = "TP2=''"
            .Formulas(30) = "TP3='X'"
            .Formulas(31) = "TP4=''"
            .Formulas(32) = "TPLIQPLAN=''"
            .Formulas(33) = "TPREGPLAN='" & frPlanAFP.xRegNum.Text & "'"
            .Formulas(34) = "INTERES=0"
        End If
        If frPlanAFP.fTipoPago(3).Value Then
            .Formulas(28) = "TP1=''"
            .Formulas(29) = "TP2=''"
            .Formulas(30) = "TP3=''"
            .Formulas(31) = "TP4='X'"
            .Formulas(32) = "TPLIQPLAN='" & frPlanAFP.xLiqNum.Text & "'"
            .Formulas(33) = "TPREGPLAN=''"
            .Formulas(34) = "INTERES=0"
        End If
        .Formulas(35) = "SREMUASEG=" & Format(SREMUASEG, "#0.00")
        .Formulas(36) = "SAPOROBLI=" & Format(SAPOROBLI, "#0.00")
'        .FORMULAS(37) = "SAPORVOLT=" & FORMAT(SAPORVOLT, "#0.00")
'        .FORMULAS(38) = "SAPORVOLE=" & FORMAT(SAPORVOLE, "#0.00")
'        .FORMULAS(39) = "SAPOREMP=" & FORMAT(SAPOREMP, "#0.00")
        .Formulas(37) = "SAPORVOLT=" & Format(0, "#0.00")
        .Formulas(38) = "SAPORVOLE=" & Format(0, "#0.00")
        .Formulas(39) = "SAPOREMP=" & Format(0, "#0.00")
        .Formulas(40) = "STOTALFONDO=" & Format(XTOTALFONDO, "#0.00")
        .Formulas(41) = "STOTALRC=" & Format(STOTALRC, "#0.00")
        .Formulas(42) = "SSEGUROS=" & Format(SSEGUROS, "#0.00")
        .Formulas(43) = "SCOMISION=" & Format(SCOMISION, "#0.00")
        .Formulas(44) = "SIM1=" & Format(S1M1, "#0.00")
        .Formulas(45) = "SIM2=" & Format(S1M2, "#0.00")
        .Formulas(46) = "STOT1=" & Format(STOT1, "#0.00")
        .Formulas(47) = "STOT2=" & Format(STOT2, "#0.00")
        .Formulas(48) = "NUMAFIL=" & frPlanAFP.cmImprimir.Tag
        .Action = 1
    End With
    Set RsEmp = Nothing
    Command2.Enabled = True
    
    Screen.MousePointer = 1
End Sub

Private Sub PAGRESTANTES()
Dim RSAUX As New ADODB.Recordset
Dim PAGREST As Integer, NTPAG As Integer
Dim I As Integer
    PAGREST = Val(xdPag.Text)
    If TOP = 0 Then
        MsgBox "Por lo menos bede imprimir un registro", vbExclamation
        xdPag.SetFocus
        Exit Sub
    End If
    
    RSAUX.Open " [##_TMPPLANAFPREST" & VGL_COMPUTER & "] ", DBSYSTEM, adOpenKeyset, adLockOptimistic
    If RSAUX.RecordCount = 0 Then
        MsgBox "No hay registros para las demas paginas"
        Exit Sub
    End If
    If PAGREST >= 19 Then
        MsgBox "Pueden ser hasta 19 registros en las demas hojas", vbInformation
        Exit Sub
    End If
    Screen.MousePointer = 11
    NTPAG = 1
    RSAUX.MoveFirst
    For I = 1 To RSAUX.RecordCount
        RSAUX!PAG = NTPAG
        RSAUX.Update
        If I Mod (PAGREST * NTPAG) = 0 Then NTPAG = NTPAG + 1
        RSAUX.MoveNext
    Next
    If ExisteTablaAux(" [##_TMPPLANAFP" & VGL_COMPUTER & "] ") Then DBSYSTEM.Execute "DROP TABLE  [##_TMPPLANAFP" & VGL_COMPUTER & "] "
    DBSYSTEM.Execute "SELECT * INTO  [##_TMPPLANAFP" & VGL_COMPUTER & "]  FROM  [##_TMPPLANAFPREST" & VGL_COMPUTER & "] "
        
    Dim RsEmp As New ADODB.Recordset
    RsEmp.Open "EMPRESA", DBSYSTEM, adOpenStatic
    With frPlanAFP.RptAFP
        .Reset
        .ReportFileName = REGSISTEMA.REPORTES & "PLAN0040.RPT"
        .Connect = "DSN=" & VGL_SERVER & ";User=sa;PWD=;DSQ=" & REGSISTEMA.BASESQL & ""
        .StoredProcParam(0) = " [##_TMPPLANAFP" & VGL_COMPUTER & "] "
        .Destination = crptToWindow
        .WindowTitle = "PLAN0040 -RESTO DE HOJAS DE PLANILLA DE PAGO DE APORTES "
        .WindowShowPrintBtn = True
        .WindowShowRefreshBtn = True
        .WindowShowSearchBtn = True
        .WindowShowPrintSetupBtn = True
        .ReplaceSelectionFormula "{ASISTMP.CODAFP}='" & frPlanAFP.RSRESAFP!Codigo & "'"
        .Formulas(0) = "NOMAFP='" & frPlanAFP.RSRESAFP!NOMAFP & "'"
        .Formulas(1) = "NOMEMPRESA='" & REGSISTEMA.EMPRESA & "'"
        .Formulas(2) = "NUMPLANILLA='" & frPlanAFP.RSRESAFP!NUMPLANILLA & "'"
        .Formulas(3) = "RUC='" & REGSISTEMA.RUC & "'"
        .Formulas(4) = "MESPAGO='" & Right(frPlanAFP.dgAFPs.Caption, 7) & "'"
        .Action = 1
    End With
    Set RsEmp = Nothing
    
    Screen.MousePointer = 1
End Sub


Private Sub Command2_Click()
    PAGRESTANTES
End Sub

Private Sub Command3_Click()
    Unload Me
End Sub

Private Sub FORM_UNLOAD(CANCEL As Integer)
On Error GoTo TemporalError
    If ExisteTablaAux("[##_TMPPLANAFP" & VGL_COMPUTER & "]") Then DBSYSTEM.Execute "DROP TABLE  [##_TMPPLANAFP" & VGL_COMPUTER & "] "
    DBSYSTEM.Execute "SELECT * INTO  [##_TMPPLANAFP" & VGL_COMPUTER & "]  FROM  [##_TMPPLANAFP2" & VGL_COMPUTER & "] "
Exit Sub
TemporalError:
End Sub


