VERSION 5.00
Object = "{C5FF36B7-A67E-11D3-9D9D-E6F193F7F854}#9.0#0"; "ApliCTxt.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form Frm_EXI 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reporte de Idiomas"
   ClientHeight    =   2640
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4650
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   4650
   Begin VB.Frame Frame1 
      Height          =   525
      Left            =   60
      TabIndex        =   8
      Top             =   1965
      Width           =   2820
      Begin VB.OptionButton Option3 
         Caption         =   "Activos"
         Height          =   285
         Left            =   150
         TabIndex        =   10
         Top             =   165
         Width           =   1050
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Todos"
         Height          =   285
         Left            =   1425
         TabIndex        =   9
         Top             =   165
         Width           =   1050
      End
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Todos"
      Height          =   225
      Left            =   3870
      TabIndex        =   6
      Top             =   510
      Width           =   750
   End
   Begin Crystal.CrystalReport CR1 
      Left            =   60
      Top             =   2205
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   60
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   2295
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   285
      Left            =   2955
      TabIndex        =   4
      Top             =   1875
      Width           =   1650
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancelar"
      Height          =   285
      Left            =   2940
      TabIndex        =   3
      Top             =   2220
      Width           =   1650
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   90
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1365
      Width           =   3675
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Empleados por Idioma"
      Height          =   285
      Left            =   30
      TabIndex        =   1
      Top             =   1005
      Width           =   3420
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Idiomas por Empleados"
      Height          =   255
      Left            =   90
      TabIndex        =   0
      Top             =   120
      Width           =   3285
   End
   Begin AplisetControlText.Aplitext xTrab 
      Height          =   285
      Left            =   75
      TabIndex        =   7
      Top             =   450
      Width           =   3630
      _ExtentX        =   6403
      _ExtentY        =   503
      Locked          =   -1  'True
      Text            =   ""
   End
End
Attribute VB_Name = "Frm_EXI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS_EMPL, RS_IDIOMA As ADODB.Recordset
Private Sub Combo1_Click()
    Combo3.ListIndex = Combo1.ListIndex
End Sub

Private Sub CHECK1_CLICK()
If Check1.Value = 1 Then
    xTrab.Text = ""
    xTrab.Tag = ""
End If
End Sub
Private Sub Command1_Click()
Dim VAR, Reporte, VAR_UNION5 As String
If REGSISTEMA.VALRRHH Then
    
        DBSTARPLAN.Execute "EXECUTE IDIOMAS_TRAB '" & REGSISTEMA.BASESQL & "', " & IIf(Option3.Value = True, 0, 1) & "," & IIf(Option1.Value = True, 0, 1) & ", " & _
                            "'" & xTrab.Tag & "', '" & Combo2.Text & "', " & Check1.Value & ""
                            'IDIOMAS_TRAB 'NATURA',1, 0,'1902893','FRAMCES',0
                            '@BASE , @OP3 , @OP1 , @TRAB , @C2 , @CHK
'            VAR = ""
'                    If Option3.Value = True Then
'                        VAR_UNION5 = "{IDIOMAS_TRABAJADOR.SITUACIÓN} >= 3"
'                    Else
'                        VAR_UNION5 = "{IDIOMAS_TRABAJADOR.SITUACIÓN} < 3"
'                    End If
            If Option1.Value = True Then
                Reporte = "PLRH0011.RPT"
'                If Check1.Value = 0 Then
'                    If Len(xTrab.Tag) > 0 Then
'                        VAR = " AND {IDIOMAS_TRABAJADOR.COD_TRAB}='" & xTrab.Tag & "'"
'                    Else
'                        MsgBox "Escoga un empleado", vbInformation
'                        xTrab.SetFocus
'                        Exit Sub
'                    End If
'                End If
            Else
                Reporte = "PLRH0012.RPT"
'                If Combo2.ListIndex > 0 Then
'                    VAR = " AND {IDIOMAS_TRABAJADOR.IDIO_DESCRIP}='" & Combo2.Text & "'"
'                End If
            End If
            With CR1
                    .Reset
                    .Connect = "DSN=" & VGL_SERVER & ";User=sa;PWD=;DSQ=" & REGSISTEMA.BASESQL & ""
                    .ReportFileName = REGSISTEMA.REPORTES & Reporte
                    .WindowTitle = Reporte & Me.Caption
                    '.SelectionFormula = VAR_UNION5 & VAR
                    .Destination = crptToWindow
                    .WindowState = crptMaximized
                    .WindowShowPrintBtn = True
                    .WindowShowRefreshBtn = True
                    .WindowShowSearchBtn = True
                    .WindowShowPrintSetupBtn = True
                    .Formulas(0) = "XEMPRESA='" & REGSISTEMA.EMPRESA & "'"
                    If .Status <> 2 Then .Action = 1
            End With
    Else
        MsgBox "UD NO TIENE PERMISO PARA EJECUTAR ESTA OPERACION", vbCritical, "ADVERTENCIA"
        Exit Sub
    End If
End Sub
Private Sub Command2_Click()
Unload Me
End Sub
Private Sub Form_Activate()
Combo2.Clear
Combo2.AddItem "TODOS LOS IDIOMAS"
RS_IDIOMA.Open "SELECT DISTINCT IDIO_DESCRIP FROM IDIOMAS", DBSYSTEM, adOpenStatic, adLockOptimistic
If RS_IDIOMA.RecordCount Then
    While Not RS_IDIOMA.EOF
        Combo2.AddItem RS_IDIOMA.Fields(0)
        RS_IDIOMA.MoveNext
    Wend
End If
    RS_IDIOMA.Close
    Combo2.ListIndex = 0
End Sub

Private Sub FORM_KEYDOWN(KEYCODE As Integer, Shift As Integer)
            If KEYCODE = vbKeyEscape Then
                Unload Me
            End If
End Sub

Private Sub Form_Load()
    Set RS_EMPL = New ADODB.Recordset
    Set RS_IDIOMA = New ADODB.Recordset
    Option1.Value = True
    Option3.Value = True
End Sub
Private Sub OPTION1_CLICK()
    xTrab.Locked = True
    Check1.Enabled = True
    Combo2.Enabled = False
End Sub
Private Sub OPTION2_Click()
    xTrab.Locked = False
    Check1.Enabled = False
    Combo2.Enabled = True
End Sub
Private Sub XTRAB_DBLCLICK()
    Dim RSTRAB As New ADODB.Recordset
    RSTRAB.Open "SELECT CODTRAB, NOMBRES FROM VWTRABAJ", DBSYSTEM, adOpenKeyset, adLockReadOnly
    If RSTRAB.EOF Or RSTRAB.RecordCount = 0 Then
        MsgBox "NO SE HAN ENCONTRADO REGISTRO DE TRABAJADORES", vbCritical
        Set RSTRAB = Nothing
        Exit Sub
    End If
    frmComun.CONECTAR RSTRAB
    frmComun.Show 1
    If VGUTIL(1) <> "" Then
        xTrab.Tag = RSTRAB!CODTRAB
        xTrab.Text = RSTRAB!CODTRAB & " : " & RSTRAB!NOMBRES
        Check1.Value = 0
    End If
    Set RSTRAB = Nothing
End Sub

