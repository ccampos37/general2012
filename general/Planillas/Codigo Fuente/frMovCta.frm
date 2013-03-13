VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C5FF36B7-A67E-11D3-9D9D-E6F193F7F854}#9.0#0"; "ApliCTxt.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frMoviCta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Movimiento de Cuentas"
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6600
   Icon            =   "frMovCta.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   6600
   StartUpPosition =   3  'Windows Default
   Begin Crystal.CrystalReport Reporte 
      Left            =   5850
      Top             =   4065
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   3518
      TabIndex        =   19
      Top             =   3990
      Width           =   1305
   End
   Begin VB.CommandButton cmAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   1778
      TabIndex        =   18
      Top             =   3990
      Width           =   1305
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos del Movimiento"
      Height          =   3675
      Left            =   90
      TabIndex        =   0
      Top             =   120
      Width           =   6225
      Begin AplisetControlText.Aplitext xDesc 
         Height          =   345
         Left            =   2010
         TabIndex        =   11
         Top             =   1080
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   609
         MaxLength       =   50
         Text            =   ""
      End
      Begin VB.ComboBox xMoneda 
         Height          =   315
         ItemData        =   "frMovCta.frx":030A
         Left            =   2010
         List            =   "frMovCta.frx":0314
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   3180
         Width           =   1665
      End
      Begin AplisetControlText.Aplitext xMeses 
         Height          =   315
         Left            =   2010
         TabIndex        =   16
         Top             =   2835
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   556
         Text            =   "0"
         Entero          =   -1  'True
         TipoDato        =   "N"
      End
      Begin MSComCtl2.DTPicker xFechaIni 
         Height          =   315
         Left            =   2010
         TabIndex        =   15
         Top             =   2490
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   556
         _Version        =   393216
         Format          =   16711681
         CurrentDate     =   36501
      End
      Begin AplisetControlText.Aplitext xPorcQ 
         Height          =   315
         Left            =   2010
         TabIndex        =   14
         Top             =   2145
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   556
         Text            =   "0.00"
         Redondear       =   -1  'True
         TipoDato        =   "N"
      End
      Begin AplisetControlText.Aplitext xInteres 
         Height          =   315
         Left            =   2010
         TabIndex        =   13
         Top             =   1785
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   556
         Text            =   "0.000"
         Redondear       =   -1  'True
         DigitRound      =   3
         TipoDato        =   "N"
      End
      Begin AplisetControlText.Aplitext xCapital 
         Height          =   315
         Left            =   2010
         TabIndex        =   12
         Top             =   1440
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   556
         Text            =   "0"
         TipoDato        =   "N"
      End
      Begin AplisetControlText.Aplitext xCodTrab 
         Height          =   345
         Left            =   2010
         TabIndex        =   10
         Top             =   707
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   609
         Locked          =   -1  'True
         Text            =   ""
      End
      Begin AplisetControlText.Aplitext xCodMov 
         Height          =   315
         Left            =   2010
         TabIndex        =   9
         Top             =   360
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   556
         Locked          =   -1  'True
         Text            =   ""
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Descripción"
         Height          =   195
         Left            =   210
         TabIndex        =   20
         Top             =   1110
         Width           =   840
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Moneda"
         Height          =   195
         Left            =   210
         TabIndex        =   8
         Top             =   3240
         Width           =   585
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Num. de Meses"
         Height          =   195
         Left            =   210
         TabIndex        =   7
         Top             =   2880
         Width           =   1110
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Fecha de Inicio"
         Height          =   195
         Left            =   210
         TabIndex        =   6
         Top             =   2535
         Width           =   1095
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Porc. en Quincena"
         Height          =   195
         Left            =   210
         TabIndex        =   5
         Top             =   2190
         Width           =   1335
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Interés"
         Height          =   195
         Left            =   210
         TabIndex        =   4
         Top             =   1830
         Width           =   480
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Capital"
         Height          =   195
         Left            =   210
         TabIndex        =   3
         Top             =   1485
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Trabajador"
         Height          =   195
         Left            =   210
         TabIndex        =   2
         Top             =   771
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         Height          =   195
         Left            =   210
         TabIndex        =   1
         Top             =   420
         Width           =   495
      End
   End
End
Attribute VB_Name = "frMoviCta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RSTRABS As New ADODB.Recordset
Dim RSMOV2 As New ADODB.Recordset

Private Sub CMACEPTAR_CLICK()
    If xCodTrab.Tag = "" Then
        MsgBox "NO HA SELECCIONADO UN TRBAJADOR PARA LA OPERACIÓN", vbCritical
        xCodTrab.SetFocus
        Exit Sub
    End If
    If xDesc.Text = "" Then
        MsgBox "DEBE TENER UNA DESCRIPCIÓN VÁLIDA", vbCritical
        xDesc.SetFocus
        Exit Sub
    End If
    If Val(xCapital.Text) <= 0 Then
        MsgBox "DEBE TENER UN MONTO VÁLIDO", vbCritical
        xCapital.SetFocus
        Exit Sub
    End If
    If xMeses.Text = "0" Then
        MsgBox "DEBE TENER UN MES VÁLIDO", vbCritical
        xMeses.SetFocus
        Exit Sub
    End If
    If xMoneda.ListIndex = -1 Then
        MsgBox "SELECCIONE LA MONEDA DE PAGO", vbCritical
        xMoneda.SetFocus
        Exit Sub
    End If
    Dim NMON As Double, NXMES As Double
    NMON = Val(xCapital.Text) + Val(xCapital.Text) * Val(xInteres.Text) / 100
    NXMES = NMON / Val(xMeses.Text)
    MsgBox "EL MONTO TOTAL A PAGAR SERÁ DE " & Format(NMON, "###,##0.00 ") & " " & xMoneda.Text & Chr(13) & Chr(10) & "PAGARÁ " & Format(NXMES, "###,##0.00 ") & " " & xMoneda.Text & " MENSUALES, CON UN DESCUENTO AL SALDO QUINCENALMENTE DE " & Format((NXMES * Val(xPorcQ.Text) / 100), "###,##0.00 ") & " " & xMoneda.Text, vbInformation
    With RSMOV2
    If VPTAREA = "NUEVO" Then
        RSMOV2.AddNew
        !CODGRUPO = VPTRASPRM
        !CAPITAL = XCAPIT
        !CODTRAB = xCodTrab.Tag
        !SALDO = Val(xCapital.Text)
    End If
    !CAPITAL = Val(xCapital.Text)
    !INTERES = Val(xInteres.Text)
    !PORCQUINC = Val(xPorcQ.Text)
    !FECHAINI = xFechaIni.Value
    !NUMMESES = Val(xMeses.Text)
    !CUOTA = NXMES
    !Moneda = xMoneda.ListIndex
    !CODTRAB = xCodTrab.Tag
    !TIPOGRUPO = frCuentas.xTipo.ListIndex + 1
    !DESCRIPCION = xDesc.Text
    !PROGRAMADO = 0
    .Update
    VPTAREA = "ACEPTÓ"
    End With
    Call ACTSALDO(VPCODTMP)
    'IMPRIMIR EL VOUCHER DEL MOVIMIENTO
    If MsgBox(" IMPRIMIR RECIBO DE MOVIMIENTO ", vbYesNo + vbQuestion) = vbYes Then
        If ExisteTablaAux(" [##TMPVOUCHER" & VGL_COMPUTER & "] ") Then DBSYSTEM.Execute "DROP TABLE  [##TMPVOUCHER" & VGL_COMPUTER & "] "
        DBSYSTEM.Execute "CREATE TABLE  [##TMPVOUCHER" & VGL_COMPUTER & "]  ([TEMP] VARCHAR(1))"
       Screen.MousePointer = 11
        With Reporte
            .Reset
            .WindowTitle = "PLAN0073.RPT - RESUMEN"
            .ReportFileName = REGSISTEMA.REPORTES & "PLAN0073.RPT"
            .Connect = "DSN=" & VGL_SERVER & ";User=sa;PWD=;DSQ=" & REGSISTEMA.BASESQL & ""
            .StoredProcParam(0) = " [##TMPVOUCHER" & VGL_COMPUTER & "] "
            .Destination = crptToWindow
            .WindowState = crptMaximized
            .WindowShowPrintBtn = True
            .WindowShowRefreshBtn = True
            .WindowShowSearchBtn = True
            .WindowShowPrintSetupBtn = True
            .Formulas(0) = "XEMPRESA='" & REGSISTEMA.EMPRESA & "'"
            .Formulas(1) = "XRUC='" & REGSISTEMA.RUC & "'"
            .Formulas(2) = "XTRAB='" & xCodTrab.Text & "'"
            .Formulas(3) = "XDESC='" & xDesc.Text & "'"
            .Formulas(4) = "XVALOR='" & Format(Val(xCapital.Text), "###,###,##0.00 ") & "'"
            .Formulas(5) = "XFECHA='" & Format(xFechaIni.Value, "DD/MM/YYYY") & "'"
            .Formulas(6) = "XMON=" & xMoneda.ListIndex
            .Formulas(7) = "XCANT='" & NUMLET(Val(xCapital.Text)) & "'"
            If .Status <> 2 Then .Action = 1
       End With
       Screen.MousePointer = 1
       cmAceptar.Enabled = False
       cmCancelar.Caption = "&SALIR"
       Else:
       Unload Me
    End If
    
    
End Sub

Private Sub CMCANCELAR_CLICK()
    Unload Me
End Sub

Private Sub Form_Activate()
    If VPTAREA = "NUEVO" Then
        xCodMov.Text = "***"
        Dim XF As Form
        Dim XPROVIENE As Byte
        XPROVIENE = 0
        For Each XF In Forms
            If UCase(XF.Name) = "FRCUENTAS" Then XPROVIENE = 1
        Next
        If XPROVIENE = 1 Then xDesc.Text = frCuentas.RSCUENTAS!NOMBRE Else xDesc.Text = VGUTIL(2)
        xFechaIni.Value = Date
        xMeses.Text = 1
        xMoneda.ListIndex = 0
        xCodTrab.SetFocus
    Else
        RSMOV2.FIND "CODMOV=" & VPCODTMP
        With RSMOV2
            xCodMov.Text = !CODMOV
            xCapital.Text = !CAPITAL
            xInteres.Text = Format(!INTERES, "0.000")
            xPorcQ.Text = Format(!PORCQUINC, "0.00")
            xFechaIni.Value = !FECHAINI
            xMeses.Text = !NUMMESES
            xMoneda.ListIndex = !Moneda
            xCodTrab.Text = !CODTRAB & " : " & VPTRASPRM
            xCodTrab.Tag = !CODTRAB
            xDesc.Text = !DESCRIPCION
        End With
    End If
End Sub

Private Sub Form_Load()
    RSTRABS.Open "SELECT * FROM VWTRABACTIVO", DBSYSTEM, adOpenStatic
    RSMOV2.Open "MOVICTA", DBSYSTEM, adOpenKeyset, adLockPessimistic
End Sub

Private Sub FORM_UNLOAD(CANCEL As Integer)
    RSTRABS.Close
    Set RSTRABS = Nothing
    RSMOV2.Close
    Set RSMOV2 = Nothing
End Sub

Private Sub XCODTRAB_DblClick()
    If VPTAREA = "EDITAR" Then Exit Sub
    frmComun.CONECTAR RSTRABS
    frmComun.Show 1
    If VGUTIL(1) <> "" Then
        xCodTrab.Tag = VGUTIL(1)
        xCodTrab.Text = VGUTIL(1) & " : " & VGUTIL(2)
    End If
End Sub

