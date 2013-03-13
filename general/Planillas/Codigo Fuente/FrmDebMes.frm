VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C5FF36B7-A67E-11D3-9D9D-E6F193F7F854}#9.0#0"; "ApliCTxt.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form FrmDebMes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Debitos Cta.Cte x Mes y Conceptos"
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6315
   Icon            =   "FrmDebMes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   6315
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox ChkConc 
      Alignment       =   1  'Right Justify
      Caption         =   "Por Concepto"
      Height          =   345
      Left            =   180
      TabIndex        =   17
      Top             =   960
      Width           =   1365
   End
   Begin VB.Frame FraConcep 
      Height          =   1275
      Left            =   90
      TabIndex        =   12
      Top             =   1005
      Width           =   3495
      Begin AplisetControlText.Aplitext xConcFin 
         Height          =   315
         Left            =   1485
         TabIndex        =   16
         Top             =   810
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   556
         Text            =   ""
      End
      Begin AplisetControlText.Aplitext xConcIni 
         Height          =   330
         Left            =   1485
         TabIndex        =   15
         Top             =   405
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   582
         Text            =   ""
      End
      Begin VB.Label Label2 
         Caption         =   "Concepto Final :"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   900
         Width           =   1275
      End
      Begin VB.Label Label1 
         Caption         =   "Concepto Inicial :"
         Height          =   195
         Left            =   105
         TabIndex        =   13
         Top             =   420
         Width           =   1290
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Opciones"
      Height          =   825
      Left            =   90
      TabIndex        =   8
      Top             =   45
      Width           =   4200
      Begin VB.OptionButton xTodos 
         Caption         =   "Todos"
         Height          =   300
         Left            =   2910
         TabIndex        =   11
         Top             =   315
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
      Begin VB.OptionButton XIngresos 
         Caption         =   "&Ingresos"
         Height          =   240
         Left            =   120
         TabIndex        =   9
         Top             =   345
         Width           =   1050
      End
   End
   Begin Crystal.CrystalReport Reporte 
      Left            =   5610
      Top             =   1125
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.TextBox SqlCad 
      Height          =   285
      Left            =   3975
      TabIndex        =   7
      Text            =   "SqlCad"
      Top             =   1200
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancelar"
      Height          =   330
      Left            =   5010
      TabIndex        =   6
      Top             =   570
      Width           =   1125
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   345
      Left            =   4995
      TabIndex        =   5
      Top             =   150
      Width           =   1140
   End
   Begin VB.Frame FraFecha 
      Height          =   1290
      Left            =   3630
      TabIndex        =   0
      Top             =   1005
      Width           =   2535
      Begin VB.CheckBox ChkMes 
         Alignment       =   1  'Right Justify
         Caption         =   "Por Mes"
         Height          =   225
         Left            =   105
         TabIndex        =   18
         Top             =   15
         Width           =   945
      End
      Begin MSComCtl2.DTPicker xFechaFin 
         Height          =   300
         Left            =   915
         TabIndex        =   1
         Top             =   750
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   393216
         Format          =   61800449
         CurrentDate     =   36691
      End
      Begin MSComCtl2.DTPicker xFechaIni 
         Height          =   300
         Left            =   915
         TabIndex        =   2
         Top             =   375
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   393216
         Format          =   61800449
         CurrentDate     =   36691
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         Height          =   195
         Left            =   165
         TabIndex        =   4
         Top             =   390
         Width           =   465
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         Height          =   195
         Left            =   150
         TabIndex        =   3
         Top             =   855
         Width           =   420
      End
   End
End
Attribute VB_Name = "FrmDebMes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Dim INTO As String
    Dim FECHAINICIO As String
    Dim FECHAFIN As String
    Dim OPCION As String
    Dim OPCFECHAS As String, PERIODOINI As String, PERIODOFIN As String
    Dim OPCCONCEPTOS As String, CONCINI As String, CONCFIN As String
    Dim X As Integer
    Dim TIPO As String
'On Error GoTo ERRADO
    If VALIDAR Then Exit Sub 'VALIDANDO EL INGRESO
    
    Screen.MousePointer = 11
    INTO = " INTO  [##TMPDEBMES" & VGL_COMPUTER & "] "
    
    OPCION = ""
    OPCCONCEPTOS = ""
    OPCFECHAS = ""
    FECHAINICIO = ""
    FECHAFIN = ""
    If ChkConc.Value = 1 Then
        OPCCONCEPTOS = " AND (CTAGRUPO.CODGRUPO>= '" & xConcIni.Tag & "' AND CTAGRUPO.CODGRUPO<= '" & xConcFin.Tag & "')"
        CONCINI = xConcIni.Text
        CONCFIN = xConcFin.Text
        Else:
          CONCINI = ""
          CONCFIN = ""
    End If
    If ChkMes.Value = 1 Then
        FECHAINICIO = Trim("'" & Month(xFechaIni) & "/01/" & Year(xFechaIni) & "'")
        FECHAFIN = Trim("'" & Month(xFechaFin) & "/01/" & Year(xFechaFin) & "'")
        OPCFECHAS = " AND NOMBOL.MES BETWEEN " & FECHAINICIO & " AND " & FECHAFIN
        PERIODOINI = Month(xFechaIni) & "/" & Year(xFechaIni)
        PERIODOFIN = Month(xFechaFin) & "/" & Year(xFechaFin)
    End If
    TIPO = ""
    If XIngresos.Value Then
        OPCION = " AND PAGOSCTA.TIPO=1"
        TIPO = "INGRESOS"
    End If
    If xEgresos.Value Then
        OPCION = " AND PAGOSCTA.TIPO=2"
        TIPO = "EGRESOS"
    End If
    If xTodos.Value Then
        OPCION = ""
        TIPO = "TODOS"
    End If
    
    SqlCad.Text = " " & _
    "   SELECT  NOMBOL.CODIGO, NOMBOL.NOMBRE, CTAGRUPO.CODGRUPO, CTAGRUPO.NOMBRE AS CTA, " & _
    "     PAGOSCTA.TIPO, TRABAJADORES.CODTRAB, " & _
    "     LTRIM(TRABAJADORES.APEPAT) + ' ' + LTRIM(TRABAJADORES.APEMAT) + " & _
    "     ' ' + LTRIM([TRABAJADORES].[NOMBRE]) AS NOMBRES, " & _
    "      PAGOSCTA.MONTO, NOMBOL.MES " & INTO & _
    " FROM NOMBOL, PAGOSCTA, TRABAJADORES, MOVICTA, CTAGRUPO " & _
    "WHERE (((NOMBOL.CODIGO) = [PAGOSCTA].[CODNOMBOL]) AND " & _
    "      ((PAGOSCTA.CODTRAB) = [TRABAJADORES].[CODTRAB]) AND " & _
    "      ((PAGOSCTA.CODMOV) = [MOVICTA].[CODMOV]) AND " & _
    "      ((MOVICTA.CODGRUPO) = [CTAGRUPO].[CODGRUPO]))" & _
    OPCFECHAS & OPCCONCEPTOS & OPCION
    
    If ExisteTablaAux(" [##TMPDEBMES" & VGL_COMPUTER & "] ") Then DBSYSTEM.Execute "DROP TABLE  [##TMPDEBMES" & VGL_COMPUTER & "] "
    DBSYSTEM.Execute SqlCad.Text, X
    
    DBSTARPLAN.Execute "EXECUTE [ASISTMP] ' [##TMPDEBMES" & VGL_COMPUTER & "] '"
    
    If X = 0 Then
        MsgBox "NO SE ENCONTRARÓN REGISTROS", vbInformation
        Screen.MousePointer = 1
        Exit Sub
    End If
    
    With Reporte
        .ReportFileName = REGSISTEMA.REPORTES & "PLAN0017.RPT"
        .WindowTitle = "PLAN0017 - DEBITOS DE CTA CTE. X CONCEPTO DE GRUPO Y MES"
        .Connect = "DSN=" & VGL_SERVERREP & ";User=sa;PWD=;DSQ=planilla_pp"
        .StoredProcParam(0) = " [##TMPDEBMES" & VGL_COMPUTER & "] "
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
        .Formulas(4) = "XCONINI='" & CONCINI & "'"
        .Formulas(5) = "XCONFIN='" & CONCFIN & "'"
        .Formulas(6) = "XTIPO='" & TIPO & "'"
        .Formulas(7) = "XHORA='" & Format(Time, "HH:MM") & "'"
        If .Status <> 2 Then .Action = 1
    End With
    Screen.MousePointer = 1
 '   Exit Sub
'ERRADO: MsgBox "POR FAVOR INTENTELO DE NUEVO"
 '       Screen.MousePointer = 1
End Sub
Private Function VALIDAR() As Boolean
VALIDAR = True
    If xFechaIni > xFechaFin Then
        MsgBox "LA FECHA DESDE NO PUEDE SER MAYOR QUE LA FECHA HASTA ", vbExclamation
        Exit Function
    End If
    If ChkConc.Value = 1 Then
        If xConcIni.Text = "" Then
            MsgBox "INGRESE EL CONCEPTO INICIO", vbExclamation
            xConcIni.SetFocus
            Exit Function
        End If
        If xConcFin.Text = "" Then
            MsgBox "INGRESE EL CONCEPTO FIN", vbExclamation
            xConcFin.SetFocus
            Exit Function
        End If
    End If
VALIDAR = False
End Function


Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub XCONCFIN_DblClick()
    Dim RS As New ADODB.Recordset
    RS.Open "CTAGRUPO", DBSYSTEM, adOpenKeyset, adLockOptimistic
    frmComun.CONECTAR RS
    frmComun.Show 1
    If VGUTIL(1) <> "" Then
        xConcFin.Text = VGUTIL(1) & " :  " & VGUTIL(2)
        xConcFin.Tag = VGUTIL(1)
    End If
End Sub

Private Sub XCONCINI_DblClick()
    Dim RS As New ADODB.Recordset
    RS.Open "CTAGRUPO", DBSYSTEM, adOpenKeyset, adLockOptimistic
    frmComun.CONECTAR RS
    frmComun.Show 1
    If VGUTIL(1) <> "" Then
        xConcIni.Text = VGUTIL(1) & " :  " & VGUTIL(2)
        xConcIni.Tag = VGUTIL(1)
    End If
End Sub


