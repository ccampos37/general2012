VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C5FF36B7-A67E-11D3-9D9D-E6F193F7F854}#9.0#0"; "ApliCTxt.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form Fr_Rep_Est 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reporte de Formacion Profesional"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5955
   Icon            =   "Fr_Rep_Est.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   5955
   Begin VB.OptionButton Option1 
      Caption         =   "Todos "
      Height          =   285
      Left            =   2640
      TabIndex        =   14
      Top             =   2400
      Width           =   1050
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Activos"
      Height          =   285
      Left            =   1500
      TabIndex        =   13
      Top             =   2385
      Width           =   1050
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Todos"
      Height          =   225
      Left            =   5100
      TabIndex        =   11
      Top             =   240
      Width           =   750
   End
   Begin Crystal.CrystalReport CR1 
      Left            =   75
      Top             =   2490
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   4200
      TabIndex        =   10
      Top             =   2700
      Width           =   1650
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   345
      Left            =   4200
      TabIndex        =   9
      Top             =   2295
      Width           =   1650
   End
   Begin VB.ComboBox Combo4 
      Height          =   315
      Left            =   1335
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1410
      Width           =   4545
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   1335
      TabIndex        =   5
      Top             =   1800
      Width           =   1560
      _ExtentX        =   2752
      _ExtentY        =   556
      _Version        =   393216
      Format          =   62259201
      CurrentDate     =   36867
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   1335
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1005
      Width           =   4545
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   1335
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   600
      Width           =   4545
   End
   Begin AplisetControlText.Aplitext xTrab 
      Height          =   285
      Left            =   1350
      TabIndex        =   12
      Top             =   195
      Width           =   3630
      _ExtentX        =   6403
      _ExtentY        =   503
      Locked          =   -1  'True
      Text            =   ""
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Centro de Estudio"
      Height          =   195
      Left            =   0
      TabIndex        =   8
      Top             =   1455
      Width           =   1260
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Desde"
      Height          =   195
      Left            =   765
      TabIndex        =   6
      Top             =   1845
      Width           =   465
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Estudio"
      Height          =   195
      Left            =   690
      TabIndex        =   4
      Top             =   1065
      Width           =   525
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Tipo Estudio"
      Height          =   195
      Left            =   315
      TabIndex        =   2
      Top             =   645
      Width           =   885
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Empleados"
      Height          =   195
      Left            =   435
      TabIndex        =   0
      Top             =   225
      Width           =   780
   End
End
Attribute VB_Name = "Fr_Rep_Est"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS_EMPL As ADODB.Recordset, RS_TIPO_EST As ADODB.Recordset, RS_EST As ADODB.Recordset, RS_CE As ADODB.Recordset

Private Sub Combo1_Click()
    Combo5.ListIndex = Combo1.ListIndex
End Sub
Private Sub CHECK1_CLICK()
If Check1.Value = 1 Then
    xTrab.Text = ""
    xTrab.Tag = ""
End If
End Sub
Private Sub Command1_Click()

If ExisteTablaAux("##TMP_ESTUDIO") Then DBSTARPLAN.Execute "DROP TABLE ##TMP_ESTUDIO"

DBSTARPLAN.Execute "EXECUTE SP_REPORTE_ESTUDIOS '" & REGSISTEMA.BASESQL & "'," & Check1.Value & ",'" & xTrab.Tag & "', " & Combo2.ListIndex & _
",'" & Combo2.Text & "'," & Combo3.ListIndex & ",'" & Combo3.Text & "'," & Combo4.ListIndex & ",'" & Combo4.Text & "'," & _
DateSQL(DTPicker1.Value) & ", " & IIf(Option1.Value = True, 0, 1) & ""

'@BASE VARCHAR(20), @CH INT, @TRAB VARCHAR(10), @C2 INT, @COMBO2 VARCHAR(50),@C3 INT,
'@COMBO3 VARCHAR(100), @C4 INT, @COMBO4 VARCHAR(50), @FECHA DATETIME, @OPT1 INT

Dim TERMINO As String, VAR_UNION5 As String, VAR_UNION1 As String, VAR_UNION2 As String, VAR_UNION3 As String, VAR_UNION4 As String

    If REGSISTEMA.VALRRHH = True Then
            With CR1
                .Reset
                .Connect = "DSN=" & VGL_SERVER & ";User=sa;PWD=;DSQ=" & REGSISTEMA.BASESQL & ""
                If Check1.Value = 0 Then
                    .WindowTitle = "PLRH0007.RPT - " & Me.Caption
                    .ReportFileName = REGSISTEMA.REPORTES & "PLRH0007.RPT"
                Else
                    .WindowTitle = "PLRH0008.RPT - " & Me.Caption
                    .ReportFileName = REGSISTEMA.REPORTES & "PLRH0008.RPT"
                End If
                .StoredProcParam(0) = "##TMP_ESTUDIO"
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
Combo3.Clear
Combo4.Clear
Combo2.AddItem "TODOS LOS TIPOS DE ESTUDIO"
RS_TIPO_EST.Open "SELECT DISTINCT TIPOEST_DESCRIP FROM ESTUDIOS", DBSYSTEM, adOpenStatic, adLockOptimistic
If RS_TIPO_EST.RecordCount Then
    While Not RS_TIPO_EST.EOF
        Combo2.AddItem RS_TIPO_EST.Fields(0)
        RS_TIPO_EST.MoveNext
    Wend
End If
RS_TIPO_EST.Close
Combo3.AddItem "TODOS LOS ESTUDIOS"
RS_EST.Open "SELECT DISTINCT EST_CARRERA FROM ESTUDIOS", DBSYSTEM, adOpenStatic, adLockOptimistic
If RS_EST.RecordCount Then
    While Not RS_EST.EOF
        Combo3.AddItem RS_EST.Fields(0)
        RS_EST.MoveNext
    Wend
End If
RS_EST.Close
Combo4.AddItem "TODOS LOS CENTROS DE ESTUDIOS"
RS_CE.Open "SELECT DISTINCT DESCESTUDIOS FROM ESTUDIOS", DBSYSTEM, adOpenStatic, adLockOptimistic
If RS_CE.RecordCount Then
    While Not RS_CE.EOF
        Combo4.AddItem RS_CE.Fields(0)
        RS_CE.MoveNext
    Wend
End If
RS_CE.Close
Combo2.ListIndex = 0
Combo3.ListIndex = 0
Combo4.ListIndex = 0
End Sub
Private Sub FORM_KEYDOWN(KEYCODE As Integer, Shift As Integer)
If KEYCODE = vbKeyEscape Then
    Unload Me
End If
End Sub
Private Sub Form_Load()
Set RS_EMPL = New ADODB.Recordset
Set RS_TIPO_EST = New ADODB.Recordset
Set RS_EST = New ADODB.Recordset
Set RS_CE = New ADODB.Recordset
Check1.Value = 1
Option2.Value = True
xTrab.Text = ""
xTrab.Tag = ""
DTPicker1.Value = #1/1/2000#
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

