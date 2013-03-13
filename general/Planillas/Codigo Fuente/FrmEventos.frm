VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C5FF36B7-A67E-11D3-9D9D-E6F193F7F854}#9.0#0"; "ApliCTxt.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form FrmEventos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Eventos"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9165
   Icon            =   "FrmEventos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   9165
   Begin VB.CommandButton Command1 
      Caption         =   "&Imprimir"
      Height          =   360
      Index           =   4
      Left            =   5310
      TabIndex        =   6
      Top             =   5610
      Width           =   960
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Grabar"
      Height          =   360
      Index           =   3
      Left            =   4095
      TabIndex        =   20
      Top             =   5610
      Width           =   960
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Eliminar"
      Height          =   360
      Index           =   2
      Left            =   2895
      TabIndex        =   5
      Top             =   5610
      Width           =   960
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Modificar"
      Height          =   360
      Index           =   1
      Left            =   1695
      TabIndex        =   4
      Top             =   5610
      Width           =   960
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Nuevo"
      Height          =   360
      Index           =   0
      Left            =   495
      TabIndex        =   3
      Top             =   5610
      Width           =   960
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cancelar"
      Height          =   360
      Index           =   5
      Left            =   6510
      TabIndex        =   22
      Top             =   5610
      Width           =   960
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Salir"
      Height          =   375
      Index           =   6
      Left            =   7710
      TabIndex        =   7
      Top             =   5610
      Width           =   960
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Mostrar"
      Height          =   315
      Left            =   5985
      TabIndex        =   41
      Top             =   525
      Width           =   870
   End
   Begin MSComCtl2.DTPicker DTPicker6 
      Height          =   300
      Left            =   4215
      TabIndex        =   1
      Top             =   540
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   529
      _Version        =   393216
      Format          =   16580609
      CurrentDate     =   36867
   End
   Begin MSComCtl2.DTPicker DTPicker5 
      Height          =   315
      Left            =   1080
      TabIndex        =   0
      Top             =   555
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      _Version        =   393216
      Format          =   16580609
      CurrentDate     =   36867
   End
   Begin AplisetControlText.Aplitext Aplitext1 
      Height          =   285
      Left            =   4920
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   180
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   503
      Locked          =   -1  'True
      Text            =   ""
   End
   Begin VB.Frame Frame2 
      Height          =   4425
      Left            =   30
      TabIndex        =   26
      Top             =   1080
      Width           =   9045
      Begin VB.CommandButton Command4 
         Caption         =   "+"
         Height          =   330
         Left            =   5385
         TabIndex        =   45
         Top             =   810
         Width           =   360
      End
      Begin VB.CommandButton Command3 
         Caption         =   "+"
         Height          =   330
         Left            =   5385
         TabIndex        =   44
         Top             =   420
         Width           =   360
      End
      Begin VB.Frame Frame5 
         Caption         =   "Asistencia"
         Height          =   660
         Left            =   6315
         TabIndex        =   42
         Top             =   720
         Width           =   2475
         Begin VB.OptionButton Option4 
            Caption         =   "No"
            Height          =   225
            Left            =   1395
            TabIndex        =   12
            Top             =   285
            Width           =   720
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Si"
            Height          =   225
            Left            =   360
            TabIndex        =   11
            Top             =   270
            Width           =   720
         End
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   2445
         TabIndex        =   9
         Top             =   810
         Width           =   2925
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   2445
         TabIndex        =   8
         Top             =   420
         Width           =   2925
      End
      Begin VB.Frame Frame4 
         Caption         =   "Fechas"
         Height          =   1245
         Left            =   1185
         TabIndex        =   34
         Top             =   2970
         Width           =   4710
         Begin VB.OptionButton Option2 
            Caption         =   "Unico"
            Height          =   285
            Left            =   180
            TabIndex        =   14
            Top             =   270
            Width           =   840
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Rango"
            Height          =   240
            Left            =   165
            TabIndex        =   16
            Top             =   735
            Width           =   885
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   285
            Left            =   2685
            TabIndex        =   15
            Top             =   255
            Width           =   1920
            _ExtentX        =   3387
            _ExtentY        =   503
            _Version        =   393216
            Format          =   16580609
            CurrentDate     =   36864
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   285
            Left            =   2685
            TabIndex        =   17
            Top             =   660
            Width           =   1920
            _ExtentX        =   3387
            _ExtentY        =   503
            _Version        =   393216
            Format          =   16580609
            CurrentDate     =   36864
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Fin"
            Height          =   255
            Left            =   1740
            TabIndex        =   36
            Top             =   705
            Width           =   420
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Inicio"
            Height          =   255
            Left            =   1755
            TabIndex        =   35
            Top             =   300
            Width           =   615
         End
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   7530
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   390
         Width           =   1245
      End
      Begin RichTextLib.RichTextBox RichTextBox1 
         Height          =   1170
         Left            =   2400
         TabIndex        =   13
         Top             =   1455
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   2064
         _Version        =   393217
         Enabled         =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"FrmEventos.frx":08CA
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Rango Horas"
         Height          =   195
         Left            =   6615
         TabIndex        =   18
         Top             =   2760
         Width           =   1440
      End
      Begin VB.Frame Frame3 
         Height          =   1245
         Left            =   6585
         TabIndex        =   31
         Top             =   2970
         Width           =   2145
         Begin MSComCtl2.DTPicker DTPicker3 
            Height          =   285
            Left            =   660
            TabIndex        =   19
            Top             =   345
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   503
            _Version        =   393216
            CustomFormat    =   "hh:mm tt"
            Format          =   16580611
            UpDown          =   -1  'True
            CurrentDate     =   36864
         End
         Begin MSComCtl2.DTPicker DTPicker4 
            Height          =   285
            Left            =   660
            TabIndex        =   21
            Top             =   750
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   503
            _Version        =   393216
            CustomFormat    =   "hh:mm tt"
            Format          =   16580611
            UpDown          =   -1  'True
            CurrentDate     =   36864
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Fin"
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   750
            Width           =   390
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Inicio"
            Height          =   255
            Left            =   105
            TabIndex        =   32
            Top             =   345
            Width           =   435
         End
      End
      Begin VB.Image Image4 
         Height          =   480
         Left            =   375
         Picture         =   "FrmEventos.frx":0938
         Top             =   3150
         Width           =   480
      End
      Begin VB.Image Image3 
         Height          =   720
         Left            =   330
         Picture         =   "FrmEventos.frx":0D7A
         Top             =   540
         Width           =   720
      End
      Begin VB.Label LblCodEv 
         Height          =   600
         Left            =   1230
         TabIndex        =   38
         Top             =   1995
         Visible         =   0   'False
         Width           =   870
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Estado"
         Height          =   195
         Left            =   6300
         TabIndex        =   30
         Top             =   450
         Width           =   495
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Evento"
         Height          =   255
         Left            =   1230
         TabIndex        =   29
         Top             =   450
         Width           =   855
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Sub Evento"
         Height          =   255
         Left            =   1230
         TabIndex        =   28
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Asunto"
         Height          =   255
         Left            =   1230
         TabIndex        =   27
         Top             =   1635
         Width           =   795
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00808080&
      Height          =   4515
      Left            =   30
      TabIndex        =   37
      Top             =   1035
      Width           =   9060
      Begin Crystal.CrystalReport CR1 
         Left            =   870
         Top             =   1740
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   4365
         Left            =   60
         TabIndex        =   2
         Top             =   105
         Width           =   8910
         _ExtentX        =   15716
         _ExtentY        =   7699
         _Version        =   393216
         FixedRows       =   0
         FixedCols       =   0
         BackColor       =   15924991
         FocusRect       =   2
         GridLines       =   0
         SelectionMode   =   1
      End
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "Eventos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   240
      Left            =   7335
      TabIndex        =   43
      Top             =   480
      Width           =   855
   End
   Begin VB.Image Image2 
      Height          =   240
      Left            =   8520
      Picture         =   "FrmEventos.frx":1834
      Top             =   315
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   8310
      Picture         =   "FrmEventos.frx":1C76
      Top             =   150
      Width           =   480
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hasta"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3600
      TabIndex        =   40
      Top             =   585
      Width           =   420
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Desde"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   165
      TabIndex        =   39
      Top             =   615
      Width           =   465
   End
   Begin VB.Label Label9 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Height          =   270
      Left            =   1050
      TabIndex        =   25
      Top             =   180
      Width           =   3780
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Trabajador"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   150
      TabIndex        =   24
      Top             =   210
      Width           =   765
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      Height          =   915
      Left            =   105
      Top             =   75
      Width           =   6840
   End
End
Attribute VB_Name = "FrmEventos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NUEVO As Boolean
Dim RS_EVE As ADODB.Recordset
Dim RS_MAYOR, RS1, RS2, RS3 As ADODB.Recordset
Dim VAR_GRIDDETALLE, CODIGO_CAT As String
Dim GOP As Boolean


Private Sub COMBO1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Command1(3).SetFocus
    KeyAscii = 0
End If
End Sub


Private Sub COMBO2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub




Private Sub COMBO3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
    KeyAscii = 0
End If
End Sub


Private Sub Command1_Click(INDEX As Integer)
Select Case INDEX
Case 0
    Call ENABLED_FR(True, False)
    Option2.Value = True
    VAR_MODO_EDIT = True
    Me.DTPicker1.Value = Date
    Me.DTPicker2.Value = Date
    Me.DTPicker3.Value = Time
    Option4.Value = True
    Me.DTPicker4.Value = Time
    DTPicker2.Enabled = False
    Check1.Value = 0
    Frame3.Enabled = False
    DTPicker3.Enabled = False
    DTPicker4.Enabled = False
    LblCodEv.Caption = ""
    Combo2.Text = ""
    Combo3.Text = ""
    RichTextBox1.Text = ""
    Combo1.ListIndex = -1
    NUEVO = True
Case 1
    VERIFICAR
    If Len(VAR_GRIDDETALLE) > 0 And MSFlexGrid1.Row > 2 Then
        NUEVO = False
        LLENA
        VAR_MODO_EDIT = True
        Call ENABLED_FR(True, False)
    End If
Case 2
    Dim SQLDELETE As String
    VERIFICAR
    If Len(VAR_GRIDDETALLE) > 0 Then
        If MsgBox("DESEA ELIMINAR EL EVENTO ?", vbYesNo, "CONFIRMACION") = vbYes Then
            SQLDELETE = "DELETE FROM EVENTOS WHERE CODIGO=" & CInt(VAR_GRIDDETALLE)
            DBSYSTEM.Execute SQLDELETE
            CARGAR
        End If
    End If
Case 3
    GRABA
    If GOP Then
        CARGAR
        Call ENABLED_FR(False, True)
        VAR_MODO_EDIT = False
    End If
Case 4
    With CR1
        .Reset
        .Connect = "DSN=" & VGL_SERVER & ";User=sa;PWD=;DSQ=" & REGSISTEMA.BASESQL & ""
        .WindowTitle = "PLRH0006.RPT - " & Me.Caption
        .ReportFileName = REGSISTEMA.REPORTES & "PLRH0005.RPT"
        .GroupSelectionFormula = "{SP_EVENTOS.CODTRAB}='" & Aplitext1.Text & "' AND {SP_EVENTOS.FEC_INI} in DateTime (" & Year(DTPicker5.Value) & ", " & Month(DTPicker5.Value) & ", " & Day(DTPicker5.Value) & ", 00, 00, 00) to DateTime (" & Year(DTPicker6.Value) & ", " & Month(DTPicker6.Value) & ", " & Day(DTPicker6.Value) & ",00, 00, 00)"
        .StoredProcParam(0) = REGSISTEMA.BASESQL
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .WindowShowPrintBtn = True
        .WindowShowRefreshBtn = True
        .WindowShowSearchBtn = True
        .WindowShowPrintSetupBtn = True
        .Formulas(0) = "EMPRESA='" & REGSISTEMA.EMPRESA & "'"
        If .Status <> 2 Then .Action = 1
    End With
Case 5
    Call ENABLED_FR(False, True)
    VAR_MODO_EDIT = False
Case 6
    Unload Me
End Select
End Sub

Private Sub Command2_Click()
CARGAR
End Sub

Private Sub CHECK1_CLICK()
If Check1.Value = 1 Then
    Frame3.Enabled = True
    DTPicker3.Enabled = True
    DTPicker4.Enabled = True
Else
    Frame3.Enabled = False
    DTPicker3.Enabled = False
    DTPicker4.Enabled = False
End If
End Sub

Private Sub CHECK1_KeyDown(KEYCODE As Integer, Shift As Integer)
If KEYCODE = 13 Then
    SendKeys "{TAB}"
    KEYCODE = 0
End If
End Sub


Private Sub Command3_Click()
    MantEven.Show 1
    Combo2.Clear
    SQLSTR = "SELECT * FROM CATEGORIA_EVENTOS ORDER BY DES_CAT"
    Set RS1 = New ADODB.Recordset
    RS1.Open SQLSTR, DBSYSTEM, adOpenStatic, adLockOptimistic
    While Not RS1.EOF
        Combo2.AddItem RS1.Fields(1)
        RS1.MoveNext
    Wend
End Sub

Private Sub Command4_Click()
    MantSCat.Show 1
    SQLSTR = "SELECT * FROM SUBCATEGORIA ORDER BY DES_SBCA"
    Set RS2 = New ADODB.Recordset
    RS2.Open SQLSTR, DBSYSTEM, adOpenStatic, adLockOptimistic
    Combo3.Clear
    If RS2.RecordCount Then
        While Not RS2.EOF
            Combo3.AddItem RS2.Fields(1)
            RS2.MoveNext
        Wend
    End If
End Sub

Private Sub DTPICKER1_KeyDown(KEYCODE As Integer, Shift As Integer)
    If KEYCODE = 13 Then
        SendKeys "{TAB}"
        KEYCODE = 0
    End If
End Sub


Private Sub DTPICKER3_KeyDown(KEYCODE As Integer, Shift As Integer)
If KEYCODE = 13 Then
    SendKeys "{TAB}"
    KEYCODE = 0
End If
End Sub


Private Sub DTPICKER4_KeyDown(KEYCODE As Integer, Shift As Integer)
If KEYCODE = 13 Then
    SendKeys "{TAB}"
    KEYCODE = 0
End If
End Sub


Private Sub Form_Activate()
Dim SQLSTR As String
If RS_EVE.State = 1 Then
    RS_EVE.Close
End If
CARGAR
End Sub
Private Sub FORM_KEYDOWN(KEYCODE As Integer, Shift As Integer)
If KEYCODE = vbKeyEscape Then
    Unload Me
End If
End Sub
Private Sub Form_Load()
Set RS_EVE = New ADODB.Recordset
Set RS1 = New ADODB.Recordset
Set RS2 = New ADODB.Recordset
Set RS3 = New ADODB.Recordset

VAR_MODO_EDIT = False
Call ENABLED_FR(False, True)
With MSFlexGrid1
    .Cols = 8
    .ColWidth(0) = 0
    .ColWidth(1) = 3000
    .ColWidth(2) = 3000
    .ColWidth(3) = 1000
    .ColWidth(4) = 1000
    .ColWidth(5) = 1000
    .ColWidth(6) = 1000
    .ColWidth(7) = 1500
End With
F1 = CDate("01/01/" & Year(Date))
Me.DTPicker5.Value = Format(F1, "DD/MM/YYYY")
Me.DTPicker6.Value = Format(Date, "DD/MM/YYYY")
        SQLSTR = "SELECT * FROM EVENTOS WHERE COD_TRABAJADOR='" & Aplitext1.Text & "'"
        RS_EVE.Open SQLSTR, DBSYSTEM, adOpenStatic, adLockOptimistic
        SQLSTR = "SELECT * FROM CATEGORIA_EVENTOS ORDER BY DES_CAT"
        Set RS1 = New ADODB.Recordset
        RS1.Open SQLSTR, DBSYSTEM, adOpenStatic, adLockOptimistic
        If RS1.RecordCount Then
            While Not RS1.EOF
                Combo2.AddItem RS1.Fields(1)
                RS1.MoveNext
            Wend
        End If
        RS1.Close
        SQLSTR = "SELECT * FROM SUBCATEGORIA ORDER BY DES_SBCA"
        Set RS2 = New ADODB.Recordset
        RS2.Open SQLSTR, DBSYSTEM, adOpenStatic, adLockOptimistic
        If RS2.RecordCount Then
            While Not RS2.EOF
                Combo3.AddItem RS2.Fields(1)
                RS2.MoveNext
            Wend
        End If
        RS2.Close
        SQLSTR = "SELECT * FROM ESTADO_EVENTO ORDER BY DES_ESTADO"
        RS3.Open SQLSTR, DBSYSTEM, adOpenStatic, adLockOptimistic
        If RS3.RecordCount Then
            While Not RS3.EOF
                Combo1.AddItem RS3.Fields(1)
                RS3.MoveNext
            Wend
        End If
        RS3.Close
End Sub
Public Sub ENABLED_FR(FLAG1 As Boolean, FLAG2 As Boolean)
        Frame2.Visible = FLAG1
        Frame1.Visible = FLAG2
        Command1(0).Enabled = FLAG2
        Command1(1).Enabled = FLAG2
        Command1(2).Enabled = FLAG2
        Command1(3).Enabled = FLAG1
        Command1(4).Enabled = FLAG2
        Command1(5).Enabled = FLAG1
        Command1(6).Enabled = FLAG2
        Command2.Enabled = FLAG2
        DTPicker5.Enabled = FLAG2
        DTPicker6.Enabled = FLAG2
End Sub
Public Sub LLENA()
Dim SQLMOD As String
Dim RS_MOD As ADODB.Recordset
Set RS_MOD = New ADODB.Recordset
    If Len(VAR_GRIDDETALLE) > 0 And IsNumeric(VAR_GRIDDETALLE) Then
        SQLMOD = "SELECT * FROM EVENTOS WHERE CODIGO=" & CInt(VAR_GRIDDETALLE)
        RS_MOD.Open SQLMOD, DBSYSTEM, adOpenStatic, adLockOptimistic
        If RS_MOD.RecordCount Then
            If Not IsNull(RS_MOD.Fields(0)) Then
                LblCodEv.Caption = RS_MOD.Fields(0)
            End If
            If Not IsNull(RS_MOD.Fields(2)) Then
                Combo2.Text = RS_MOD.Fields(2)
            End If
            If Not IsNull(RS_MOD.Fields(3)) Then
                Combo3.Text = RS_MOD.Fields(3)
            End If
            If Len(RS_MOD.Fields(4)) > 0 Then
                Combo1.Text = RS_MOD.Fields(4)
            Else
                Combo1.ListIndex = -1
            End If
            If Len(RS_MOD.Fields(5)) > 0 Then
                DTPicker1.Value = RS_MOD.Fields(5)
            End If
            If Len(RS_MOD.Fields(6)) > 0 Then
                DTPicker2.Value = RS_MOD.Fields(6)
                Option1.Value = True
            Else
                Option2.Value = True
                DTPicker2.Enabled = False
            End If
            If Len(RS_MOD.Fields(7)) > 0 Then
                DTPicker3.Value = RS_MOD.Fields(7)
            Else
                Check1.Value = 0
                DTPicker3.Enabled = False
                DTPicker4.Enabled = False
            End If
            If Len(RS_MOD.Fields(8)) > 0 Then
                DTPicker4.Value = RS_MOD.Fields(8)
            End If
            If Len(Trim(RS_MOD.Fields(9))) > 0 Then
                RichTextBox1.Text = RS_MOD.Fields(9)
            End If
        End If
        If RS_MOD.Fields(10) = True Then
            Me.Option3.Value = True
        Else
            Me.Option4.Value = True
        End If
        RS_MOD.Close
    End If
End Sub
Private Sub FORM_UNLOAD(CANCEL As Integer)
    VAR_SHOW = 0
End Sub

Private Sub MSFLEXGRID1_Click()
MSFlexGrid1.COL = 1
End Sub
Private Sub MSFLEXGRID1_DblClick()
Dim X, POSICIONINICIAL As Integer
        POSICIONINICIAL = MSFlexGrid1.Row
        MSFlexGrid1.Row = POSICIONINICIAL
        MSFlexGrid1.COL = 0
            If MSFlexGrid1.CellBackColor = &H808080 Then
                If MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1) = "-" Then
                    MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1) = "+"
                    For X = POSICIONINICIAL + 1 To MSFlexGrid1.Rows - 1
                        MSFlexGrid1.Row = X
                        MSFlexGrid1.COL = 0
                        If MSFlexGrid1.CellBackColor <> &H808080 Then
                            MSFlexGrid1.RowHeight(MSFlexGrid1.Row) = 0
                        Else
                            MSFlexGrid1.Row = POSICIONINICIAL
                            Exit For
                        End If
                    Next
                Else
                    MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1) = "-"
                    For X = POSICIONINICIAL + 1 To MSFlexGrid1.Rows - 1
                        MSFlexGrid1.Row = X
                        MSFlexGrid1.COL = 0
                        If MSFlexGrid1.CellBackColor <> &H808080 Then
                            MSFlexGrid1.RowHeight(MSFlexGrid1.Row) = 220
                        Else
                            MSFlexGrid1.Row = POSICIONINICIAL
                            Exit For
                        End If
                    Next
                End If
            Else
                VERIFICAR
                If Len(VAR_GRIDDETALLE) > 0 And IsNumeric(VAR_GRIDDETALLE) Then
                    NUEVO = False
                    LLENA
                    Call ENABLED_FR(True, False)
                    VAR_MODO_EDIT = True
                End If
            End If
            MSFlexGrid1.Row = POSICIONINICIAL
End Sub
Private Sub MSFLEXGRID1_GOTFOCUS()
MSFlexGrid1.COL = 1
End Sub
Private Sub MSFLEXGRID1_KeyDown(KEYCODE As Integer, Shift As Integer)
If KEYCODE = vbKeyDelete Then
    Dim SQLDELETE As String
    VERIFICAR
    If Len(VAR_GRIDDETALLE) > 0 Then
        If MsgBox("Desea eliminar el evento ?", vbYesNo, "Confirmación") = vbYes Then
            SQLDELETE = "DELETE FROM EVENTOS WHERE CODIGO=" & CInt(VAR_GRIDDETALLE)
            DBSYSTEM.Execute SQLDELETE
            CARGAR
        End If
    End If
End If
End Sub
Private Sub OPTION1_CLICK()
DTPicker2.Enabled = True
End Sub
Private Sub OPTION2_Click()
DTPicker2.Enabled = False
End Sub
Public Sub CARGAR()
Set RS_MAYOR = New ADODB.Recordset
RS_MAYOR.Open "SELECT DISTINCT CATEGORIA FROM EVENTOS WHERE COD_TRABAJADOR='" & Aplitext1.Text & "' AND FEC_INI>=" & DateSQL(DTPicker5.Value) & " AND FEC_INI <=" & DateSQL(DTPicker6.Value) & "", DBSYSTEM, adOpenStatic, adLockOptimistic
With MSFlexGrid1
Dim X As Integer
    .Cols = 8
    .Rows = 3
    For X = 0 To .Cols - 1
        .TextMatrix(1, X) = "EVENTOS"
        .Row = 1
        .COL = X
        .CellAlignment = 4
        .CellFontBold = True
        .CellFontSize = 10
    Next
    .RowHeight(1) = 250
    .MergeCells = flexMergeFree
    .MergeRow(1) = True
    If RS_MAYOR.RecordCount Then
        While Not RS_MAYOR.EOF
            .Rows = MSFlexGrid1.Rows + 1
            For X = 0 To .Cols - 1
                .TextMatrix(MSFlexGrid1.Rows - 1, X) = UCase(RS_MAYOR.Fields(0))
                .Row = .Rows - 1
                .COL = X
                .CellBackColor = &H808080
                .CellFontBold = True
                .CellAlignment = 4
                .CellForeColor = vbWhite
            Next
                .TextMatrix(.Rows - 1, 1) = "-"
                .Row = .Rows - 1
                .COL = 1
                .CellFontBold = True
                .MergeRow(.Rows - 1) = True
                .MergeRow(.Rows - 1) = True
                    Set RS1 = New ADODB.Recordset
                    RS1.Open "SELECT * FROM EVENTOS WHERE COD_TRABAJADOR='" & Aplitext1.Text & "' AND CATEGORIA='" & RS_MAYOR.Fields(0) & "' AND FEC_INI>=" & DateSQL(DTPicker5.Value) & " AND FEC_INI<=" & DateSQL(DTPicker6.Value) & " ORDER BY FEC_INI", DBSYSTEM, adOpenStatic, adLockOptimistic
                    If RS1.RecordCount Then
                        While Not RS1.EOF
                            MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
                            MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 0) = UCase(RS1.Fields(0))
                            MSFlexGrid1.Row = MSFlexGrid1.Rows - 1
                            MSFlexGrid1.COL = 0
                            MSFlexGrid1.CellAlignment = 2
                            MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 1) = UCase(RS1.Fields(3))
                            MSFlexGrid1.Row = MSFlexGrid1.Rows - 1
                            MSFlexGrid1.COL = 1
                            MSFlexGrid1.CellAlignment = 2
                            If Not IsNull(RS1.Fields(9)) Then
                                MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 2) = UCase(RS1.Fields(9))
                                MSFlexGrid1.Row = MSFlexGrid1.Rows - 1
                                MSFlexGrid1.COL = 2
                                MSFlexGrid1.CellAlignment = 2
                            End If
                            If Not IsNull(RS1.Fields(5)) Then
                                MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 3) = UCase(RS1.Fields(5))
                                MSFlexGrid1.Row = MSFlexGrid1.Rows - 1
                                MSFlexGrid1.COL = 3
                                MSFlexGrid1.CellAlignment = 2
                            End If
                            If Not IsNull(RS1.Fields(6)) Then
                                MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 4) = UCase(RS1.Fields(6))
                                MSFlexGrid1.Row = MSFlexGrid1.Rows - 1
                                MSFlexGrid1.COL = 4
                                MSFlexGrid1.CellAlignment = 2
                            End If
                            If Not IsNull(RS1.Fields(7)) Then
                                MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 5) = UCase(RS1.Fields(7))
                                MSFlexGrid1.Row = MSFlexGrid1.Rows - 1
                                MSFlexGrid1.COL = 5
                                MSFlexGrid1.CellAlignment = 2
                            End If
                            If Not IsNull(RS1.Fields(8)) Then
                                MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 6) = UCase(RS1.Fields(8))
                                MSFlexGrid1.Row = MSFlexGrid1.Rows - 1
                                MSFlexGrid1.COL = 6
                                MSFlexGrid1.CellAlignment = 2
                            End If
                            MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 7) = UCase(RS1.Fields(4))
                            MSFlexGrid1.Row = MSFlexGrid1.Rows - 1
                            MSFlexGrid1.COL = 7
                            MSFlexGrid1.CellAlignment = 2
                            RS1.MoveNext
                        Wend
                    End If
                    RS1.Close
                RS_MAYOR.MoveNext
        Wend
    End If
End With
RS_MAYOR.Close
Dim SQLEXPORT, SQLDELETE As String
If ExisteTablaSQL("TMP_EVENTOS", DBSYSTEM) Then DBSYSTEM.Execute "DROP TABLE TMP_EVENTOS"
    SLQDELETE = "CREATE TABLE TMP_EVENTOS (COD_TRAB VARCHAR(10), CATEGORIA VARCHAR(100), SUBCATEGORIA VARCHAR(100), ESTADO VARCHAR(30), FI VARCHAR(15), FF VARCHAR(15), HI VARCHAR(15), HF VARCHAR(15), ASUNTO VARCHAR(800))"
    DBSYSTEM.Execute SLQDELETE 'BORRAR TODOS LOS REGISTROS
SQLEXPORT = "INSERT INTO TMP_EVENTOS (COD_TRAB, CATEGORIA, SUBCATEGORIA, ESTADO, FI, FF, HI, HF, ASUNTO) " & _
" SELECT TRABAJADORES.CODTRAB, EVENTOS.CATEGORIA, EVENTOS.SUBCATEGORIA, EVENTOS.ESTADO, EVENTOS.FEC_INI, EVENTOS.FEC_FIN, EVENTOS.HOR_FIN, EVENTOS.HOR_FIN, EVENTOS.ASUNTO " & _
" FROM TRABAJADORES INNER JOIN EVENTOS ON TRABAJADORES.CODTRAB = EVENTOS.COD_TRABAJADOR " & _
" WHERE EVENTOS.FEC_INI BETWEEN " & DateSQL(DTPicker5.Value) & " AND " & DateSQL(DTPicker6.Value) & "" & _
" ORDER BY EVENTOS.FEC_INI;"
DBSYSTEM.Execute SQLEXPORT 'INSERTAR LOS NUEVOS REGISTROS
End Sub
Public Sub GRABA()
Dim SQLEXEC As String
Dim Codigo As Integer
Dim F1, F2 As Variant
Dim H1, H2 As Variant
Codigo = GENERA
GOP = False
If Option2.Value = True Then
    F1 = DTPicker1.Value
    F2 = Null
Else
    F1 = DTPicker1.Value
    F2 = DTPicker2.Value
End If
If Check1.Value = 1 Then
    H1 = Format(DTPicker3.Value, "HH:MM:SS")
    H2 = Format(DTPicker4.Value, "HH:MM:SS")
Else
    H1 = Null
    H2 = Null
End If
Dim XVAR As Double
    If Me.Option3.Value = True Then
        XVAR = -1
    ElseIf Me.Option4.Value = True Then
        XVAR = 0
    End If
    If Len(Trim(Me.Combo2.Text)) = 0 Then
        MsgBox "Datos incompletos ingrese el evento del trabajador", vbInformation, "Información"
        Combo1.SetFocus
        Exit Sub
    End If
    If NUEVO Then
        If IsNull(F2) Then
            SQLEXEC = "INSERT INTO EVENTOS VALUES (" & Codigo & ",'" & Aplitext1.Text & "','" & Combo2.Text & "', '" & Combo3.Text & "','" & _
                Combo1.Text & "', " & DateSQL(F1) & ", NULL, '" & H1 & "','" & H2 & "', '" & RichTextBox1.Text & "'," & XVAR & " )"
        Else
            SQLEXEC = "INSERT INTO EVENTOS VALUES (" & Codigo & ",'" & Aplitext1.Text & "','" & Combo2.Text & "', '" & Combo3.Text & "','" & _
                Combo1.Text & "'," & DateSQL(F1) & "," & DateSQL(F2) & ", '" & H1 & "','" & H2 & "', '" & RichTextBox1.Text & "'," & XVAR & " )"
        End If
    Else
        SQLEXEC = "UPDATE EVENTOS SET CATEGORIA='" & Combo2.Text & "', SUBCATEGORIA='" & Combo3.Text & "', ESTADO='" & Combo1.Text & "', FEC_INI=" & DateSQL(F1) & ", FEC_FIN=" & IIf(IsNull(F2), "NULL", DateSQL(IIf(IsNull(F2), Date, F2))) & ", HOR_INI='" & H1 & "', HOR_FIN='" & H2 & "', ASUNTO='" & RichTextBox1.Text & "', INFAS=" & XVAR & " WHERE CODIGO=" & CInt(LblCodEv.Caption)
    End If
DBSYSTEM.Execute SQLEXEC
GOP = True
End Sub
Function GENERA() As String
Dim SSQL As String
Dim RS_AUX As ADODB.Recordset
SSQL = "SELECT * FROM EVENTOS ORDER BY CODIGO"
Set RS_AUX = New ADODB.Recordset
RS_AUX.Open SSQL, DBSYSTEM, adOpenStatic, adLockOptimistic
If RS_AUX.RecordCount Then
    RS_AUX.MoveLast
    GENERA = RS_AUX.Fields(0) + 1
Else
    GENERA = 1
End If
RS_AUX.Close
End Function
Public Sub VERIFICAR()
VAR_GRIDDETALLE = ""
    If Len(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0)) > 0 And MSFlexGrid1.CellBackColor <> &H808000 Then
        MSFlexGrid1.Row = MSFlexGrid1.Row
        MSFlexGrid1.COL = 0
        If MSFlexGrid1.CellFontSize <> 10 Then
            VAR_GRIDDETALLE = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0)
        End If
    End If
    MSFlexGrid1.Row = POSICIONINICIAL
End Sub

Private Sub RICHTEXTBOX1_KeyDown(KEYCODE As Integer, Shift As Integer)
    If KEYCODE = 13 Then
        SendKeys "{TAB}"
        KEYCODE = 0
    End If
End Sub

