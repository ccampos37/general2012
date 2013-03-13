VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C5FF36B7-A67E-11D3-9D9D-E6F193F7F854}#9.0#0"; "ApliCTxt.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form Fr_Rep_Eve 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reporte de Eventos"
   ClientHeight    =   2490
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5955
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2490
   ScaleWidth      =   5955
   Begin VB.OptionButton Option2 
      Caption         =   "Cesantes"
      Height          =   285
      Left            =   2535
      TabIndex        =   14
      Top             =   1800
      Width           =   1050
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Activos"
      Height          =   285
      Left            =   1275
      TabIndex        =   13
      Top             =   1785
      Width           =   1050
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Todos"
      Height          =   225
      Left            =   5085
      TabIndex        =   11
      Top             =   180
      Width           =   750
   End
   Begin Crystal.CrystalReport CR1 
      Left            =   90
      Top             =   1950
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancelar"
      Height          =   285
      Left            =   4230
      TabIndex        =   10
      Top             =   2130
      Width           =   1650
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   285
      Left            =   4230
      TabIndex        =   9
      Top             =   1785
      Width           =   1650
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   300
      Left            =   4155
      TabIndex        =   6
      Top             =   1365
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   529
      _Version        =   393216
      Format          =   62062593
      CurrentDate     =   36867
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   1305
      TabIndex        =   5
      Top             =   1350
      Width           =   1560
      _ExtentX        =   2752
      _ExtentY        =   556
      _Version        =   393216
      Format          =   62062593
      CurrentDate     =   36867
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   945
      Width           =   4545
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   525
      Width           =   4545
   End
   Begin AplisetControlText.Aplitext xTrab 
      Height          =   285
      Left            =   1335
      TabIndex        =   12
      Top             =   135
      Width           =   3630
      _ExtentX        =   6403
      _ExtentY        =   503
      Locked          =   -1  'True
      Text            =   ""
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Hasta"
      Height          =   255
      Left            =   3000
      TabIndex        =   8
      Top             =   1395
      Width           =   1035
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Desde"
      Height          =   255
      Left            =   165
      TabIndex        =   7
      Top             =   1365
      Width           =   1035
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Sub Estudio"
      Height          =   255
      Left            =   180
      TabIndex        =   4
      Top             =   1005
      Width           =   1035
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Eventos"
      Height          =   255
      Left            =   150
      TabIndex        =   2
      Top             =   600
      Width           =   1035
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Empleados"
      Height          =   255
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   1035
   End
End
Attribute VB_Name = "Fr_Rep_Eve"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS_EMPL, RS_TIPO_EST, RS_EST As ADODB.Recordset
Private Sub Combo1_Click()
    Combo4.ListIndex = Combo1.ListIndex
End Sub
Private Sub CHECK1_CLICK()
    If Check1.Value = 1 Then
        xTrab.Text = ""
    End If
End Sub
Private Sub Command1_Click()
If REGSISTEMA.VALRRHH = True Then

    If ExisteTablaAux(" [##TMPEVENTOS" & VGL_COMPUTER & "] ") Then DBSTARPLAN.Execute "DROP TABLE  [##TMPEVENTOS" & VGL_COMPUTER & "] "
    
    DBSTARPLAN.Execute "EXECUTE SP_EVENTOS_TRAB '" & REGSISTEMA.BASESQL & "', " & Combo2.ListIndex & ", '" & _
                        Combo2.Text & "'," & Combo3.ListIndex & ",'" & Combo3.Text & "'," & IIf(Option1.Value = True, 0, 1) & _
                        ", '" & xTrab.Tag & "', " & DateSQL(DTPicker1.Value) & ", " & Check1.Value & ""
                With CR1
                    .Reset
                    .Connect = "DSN=" & VGL_SERVER & ";User=sa;PWD=;DSQ=" & REGSISTEMA.BASESQL & ""
                    If Check1.Value = 0 Then
                        .WindowTitle = "PLRH0009.RPT - " & Me.Caption
                        .ReportFileName = REGSISTEMA.REPORTES & "PLRH0009.RPT"
                    Else
                        .WindowTitle = "PLRH0010.RPT - " & Me.Caption
                        .ReportFileName = REGSISTEMA.REPORTES & "PLRH0010.RPT"
                    End If
                    .StoredProcParam(0) = " [##TMPEVENTOS" & VGL_COMPUTER & "] "
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
        MsgBox "UD NO TIENE PERMISO PARA EJECUTAR ESTA OPERACION", vbCritical, "ADVERETENCIA"
        Exit Sub
    End If
End Sub
Private Sub Command2_Click()
Unload Me
End Sub
Private Sub Form_Activate()
Combo2.Clear
Combo3.Clear
Combo2.AddItem "TODOS LOS EVENTOS"
RS_TIPO_EST.Open "SELECT DISTINCT CATEGORIA FROM EVENTOS", DBSYSTEM, adOpenStatic, adLockOptimistic
If RS_TIPO_EST.RecordCount Then
    While Not RS_TIPO_EST.EOF
        Combo2.AddItem RS_TIPO_EST.Fields(0)
        RS_TIPO_EST.MoveNext
    Wend
End If
RS_TIPO_EST.Close
Combo3.AddItem "TODOS LOS SUB EVENTOS"
RS_EST.Open "SELECT DISTINCT SUBCATEGORIA FROM EVENTOS", DBSYSTEM, adOpenStatic, adLockOptimistic
If RS_EST.RecordCount Then
    While Not RS_EST.EOF
        If Len(RS_EST.Fields(0)) > 0 Then
            Combo3.AddItem RS_EST.Fields(0)
        End If
        RS_EST.MoveNext
    Wend
End If
RS_EST.Close
Combo2.ListIndex = 0
Combo3.ListIndex = 0
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
DTPicker1.Value = #1/1/2000#
DTPicker2.Value = #12/31/2000#
Check1.Value = 1
Option1.Value = True
xTrab.Text = ""
xTrab.Tag = ""
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

