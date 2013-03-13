VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRepVenAgr 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ventas por Agrupaciones"
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7905
   Icon            =   "FrmRepVenAgr.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   7905
   Begin VB.Frame Frame4 
      Height          =   1455
      Left            =   5520
      TabIndex        =   26
      Top             =   2050
      Width           =   2175
      Begin VB.OptionButton Option2 
         Caption         =   "Reporte Resumido"
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   28
         Top             =   720
         Width           =   1815
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Reporte Detallado"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   27
         Top             =   360
         Value           =   -1  'True
         Width           =   1695
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   675
      Left            =   2640
      Picture         =   "FrmRepVenAgr.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3600
      Width           =   775
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   675
      Left            =   4320
      Picture         =   "FrmRepVenAgr.frx":0D0C
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3600
      Width           =   775
   End
   Begin VB.Frame Frame3 
      Height          =   1440
      Left            =   120
      TabIndex        =   21
      Top             =   2055
      Width           =   5295
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   885
         Width           =   3735
      End
      Begin VB.TextBox TxArt1 
         Height          =   285
         Left            =   1080
         TabIndex        =   7
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox TxArt2 
         Height          =   285
         Left            =   3600
         TabIndex        =   8
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Almacen  :"
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   885
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "Articulo  "
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "Desde   :"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "Hasta  :"
         Height          =   255
         Left            =   2880
         TabIndex        =   22
         Top             =   480
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1935
      Left            =   5520
      TabIndex        =   20
      Top             =   120
      Width           =   2175
      Begin VB.OptionButton Option1 
         Caption         =   "Por Familia"
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   15
         Top             =   1320
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Por Linea"
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   14
         Top             =   960
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Por Grupo"
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   13
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Por Artículo"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   12
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Parámetros"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   5295
      Begin VB.TextBox TxLinea 
         Height          =   285
         Left            =   3720
         TabIndex        =   6
         Top             =   1500
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox TxFamilia 
         Height          =   285
         Left            =   3720
         TabIndex        =   5
         Top             =   1110
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.ComboBox CmbPV2 
         Height          =   315
         Left            =   4305
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   330
         Width           =   855
      End
      Begin VB.ComboBox CmbPV1 
         Height          =   315
         Left            =   1830
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   315
         Width           =   855
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "FrmRepVenAgr.frx":114E
         Left            =   1320
         List            =   "FrmRepVenAgr.frx":1158
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1440
         Width           =   855
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   270
         Left            =   1320
         TabIndex        =   2
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   476
         _Version        =   393216
         Format          =   24838145
         CurrentDate     =   36621
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   270
         Left            =   1320
         TabIndex        =   3
         Top             =   1080
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   476
         _Version        =   393216
         Format          =   24838145
         CurrentDate     =   36621
      End
      Begin VB.Label Label11 
         Caption         =   "Línea  :"
         Height          =   225
         Left            =   3000
         TabIndex        =   32
         Top             =   1545
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label10 
         Caption         =   "Familia :"
         Height          =   225
         Left            =   3000
         TabIndex        =   31
         Top             =   1155
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Pto. Vta. Final    :"
         Height          =   255
         Left            =   2985
         TabIndex        =   30
         Top             =   375
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Pto. Vta. Inicial   :"
         Height          =   255
         Left            =   270
         TabIndex        =   29
         Top             =   375
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Del Dia      :"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Al Dia        :"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Moneda    :"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   1440
         Width           =   855
      End
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   5400
      Top             =   3600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
End
Attribute VB_Name = "FrmRepVenAgr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cR As ADODB.Recordset
Dim cR1 As ADODB.Recordset
Dim nSw As Integer, nSw2 As Integer
Private Sub CmbPV1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub CmbPV2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub Command1_Click()
If Option1(2).Value Then
    If Trim(TxFamilia) = "" Then
        MsgBox "Ingrese Familia", vbInformation, "Información"
        TxFamilia.SetFocus: Exit Sub
    End If
ElseIf Option1(1).Value Then
    If Trim(TxFamilia) = "" Then
        MsgBox "Ingrese Familia", vbInformation, "Información"
        TxFamilia.SetFocus: Exit Sub
    ElseIf Trim(TxLinea) = "" Then
        MsgBox "Ingrese Linea", vbInformation, "Información"
        TxLinea.SetFocus: Exit Sub
    End If
End If

If DTPicker1 > DTPicker2 Then
    MsgBox "La Fecha Inicial debe ser  menor a la Fecha Final", vbInformation, "Sistema de Ventas"
    DTPicker1.SetFocus: Exit Sub
End If


If Trim(TxArt2) = "" Then TxArt2 = TxArt1

Call Rep_Opc(TxArt1, TxArt2, Mid(Combo3.text, 1, 2), DTPicker1, DTPicker2)

cConexAux.Execute "Select * From Vent_Cli"
If Option1(0).Value Then
    If Option2(0).Value Then
        CrystalReport1.ReportFileName = cRutP & "\venagr01.Rpt"
    Else
        CrystalReport1.ReportFileName = cRutP & "\venagr02.Rpt"
    End If
    CrystalReport1.SelectionFormula = "{PUNTO_VENTA.PV_EMPRESA} = '" & vGEmpresa & "'"
     CrystalReport1.Formulas(8) = ""
     CrystalReport1.Formulas(9) = ""
ElseIf Option1(3).Value Then
    CrystalReport1.ReportFileName = cRutP & "\venagf01.Rpt"
    CrystalReport1.SelectionFormula = "{PUNTO_VENTA.PV_EMPRESA} = '" & vGEmpresa & "' AND {VENT_CLI.VENFAMILIA} >= '" & TxArt1 & "' AND {VENT_CLI.VENFAMILIA} <= '" & TxArt2 & "'"
    CrystalReport1.Formulas(8) = ""
    CrystalReport1.Formulas(9) = ""
ElseIf Option1(2).Value Then
    CrystalReport1.ReportFileName = cRutP & "\venagl01.Rpt"
    CrystalReport1.SelectionFormula = "{PUNTO_VENTA.PV_EMPRESA} = '" & vGEmpresa & "' AND {VENT_CLI.VENMODELO} >= '" & TxArt1 & "' AND {VENT_CLI.VENMODELO} <= '" & TxArt2 & "' and {LINEAS.FAM_CODIGO} = '" & TxFamilia & "'"
    CrystalReport1.Formulas(8) = "Familia = '" & TxFamilia & "'"
    CrystalReport1.Formulas(9) = ""
ElseIf Option1(1).Value Then
    CrystalReport1.ReportFileName = cRutP & "\venagg01.Rpt"
    CrystalReport1.SelectionFormula = "{PUNTO_VENTA.PV_EMPRESA} = '" & vGEmpresa & "' AND {VENT_CLI.VENGRUPO} >= '" & TxArt1 & "' AND {VENT_CLI.VENGRUPO} <= '" & TxArt2 & "' and {GRUPO.FAM_CODIGO} = '" & TxFamilia & "' and {GRUPO.LIN_CODIGO} = '" & TxLinea & "'"
    CrystalReport1.Formulas(8) = "Familia = '" & TxFamilia & "'"
    CrystalReport1.Formulas(9) = "Linea = '" & TxLinea & "'"
End If

Call Ubi_Tab(CrystalReport1)
CrystalReport1.Formulas(0) = "Hora = '" & Time & "'"
CrystalReport1.Formulas(1) = "Empresa = '" & vGNomEmp & "'"
CrystalReport1.Formulas(2) = "FecIni = '" & DTPicker1 & "'"
CrystalReport1.Formulas(3) = "FecFin = '" & DTPicker2 & "'"
CrystalReport1.Formulas(4) = "ArtIni = '" & TxArt1 & "'"
CrystalReport1.Formulas(5) = "ArtFin = '" & TxArt2 & "'"
CrystalReport1.Formulas(6) = "Almacen = '" & Combo3.text & "'"
If Combo2.ListIndex = 0 Then
    CrystalReport1.Formulas(7) = "Moneda = 'Nacional'"
Else
    CrystalReport1.Formulas(7) = "Moneda = 'Extranjera'"
End If

CrystalReport1.DiscardSavedData = True
CrystalReport1.Destination = crptToWindow
If CrystalReport1.Status <> 2 Then CrystalReport1.Action = 1
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
CentForm Me
AgreCom
If CmbPV1.ListCount > 0 Then CmbPV1.ListIndex = 0
If CmbPV2.ListCount > 0 Then CmbPV2.ListIndex = 0
If Combo2.ListCount > 0 Then Combo2.ListIndex = 0
If Combo3.ListCount > 0 Then Combo3.ListIndex = 0
DTPicker1 = Date
DTPicker2 = Date
Limp_Tex
End Sub
Private Sub AgreCom()
Dim cS As String

CmbPV1.Clear
CmbPV2.Clear

cS = "Select PV_COD From PUNTO_VENTA WHERE PV_EMPRESA = '" & vGEmpresa & "' order by PV_COD"
Set cR = New ADODB.Recordset
cR.Open cS, cConexConf, adOpenStatic
Do While Not cR.EOF
    CmbPV1.AddItem cR("PV_COD")
    CmbPV2.AddItem cR("PV_COD")
    cR.MoveNext
    If cR.EOF Then Exit Do
Loop

cS = "Select Taalma,Tadescri from TabAlm order by Taalma "
Set cR1 = New ADODB.Recordset
cR1.Open cS, cconexcom, adOpenStatic
Do While Not cR1.EOF
    Combo3.AddItem cR1("Taalma") & "  " & cR1("Tadescri")
    cR1.MoveNext
    If cR1.EOF Then Exit Do
Loop
End Sub

Private Sub Option1_Click(Index As Integer)
Select Case Index
Case 0:
        Label8 = "Articulo"
        TxArt1.Enabled = True
        TxArt2.Enabled = True
        Frame4.Enabled = True
        Label10.Visible = False
        Label11.Visible = False
        TxFamilia.Visible = False
        TxLinea.Visible = False
        TxFamilia = "": TxLinea = ""
Case 1:
        Label8 = "Grupo"
        TxArt1.Enabled = True
        TxArt2.Enabled = True
        Label10.Visible = True
        Label11.Visible = True
        TxFamilia.Visible = True
        TxLinea.Visible = True
        Frame4.Enabled = False
Case 2:
        Label8 = "Linea"
        TxArt1.Enabled = True
        TxArt2.Enabled = True
        Label10.Visible = True
        Label11.Visible = False
        TxFamilia.Visible = True
        TxLinea.Visible = False
        Frame4.Enabled = False
Case 3:
        Label8 = "Familia"
        TxArt1.Enabled = True
        TxArt2.Enabled = True
        Label10.Visible = False
        Label11.Visible = False
        TxFamilia.Visible = False
        TxLinea.Visible = False
        Frame4.Enabled = False
End Select
Limp_Tex
End Sub


Private Sub TxArt1_DblClick()
Dim Adodc3 As ADODB.Recordset
Set Adodc3 = New ADODB.Recordset
If Option1(0).Value Then
    Adodc3.Open "SELECT Acodigo,adescri  FROM MAeArt", cconexcom, adOpenStatic, adLockOptimistic
    frmReferencia.Conectar Adodc3, "SELECT Acodigo,adescri  FROM MAeArt"
    frmReferencia.Label1.Caption = "Articulos"
    frmReferencia.Show vbModal
    Adodc3.Close
    If vGUtil(1) <> "" Then TxArt1 = (vGUtil(1))
ElseIf Option1(3).Value Then
    Adodc3.Open "SELECT FAM_CODIGO,FAM_NOMBRE  FROM FAMILIA", cconexcom, adOpenStatic, adLockOptimistic
    frmReferencia.Conectar Adodc3, "SELECT FAM_CODIGO,FAM_NOMBRE  FROM FAMILIA"
    frmReferencia.Label1.Caption = "Familias"
    frmReferencia.Show vbModal
    Adodc3.Close
    If vGUtil(1) <> "" Then TxArt1 = (vGUtil(1))
ElseIf Option1(2).Value Then
    Adodc3.Open "SELECT LIN_CODIGO,LIN_NOMBRE  FROM LINEAS WHERE FAM_CODIGO = '" & TxFamilia & "'", cconexcom, adOpenStatic, adLockOptimistic
    frmReferencia.Conectar Adodc3, "SELECT LIN_CODIGO,LIN_NOMBRE  FROM LINEAS WHERE FAM_CODIGO = '" & TxFamilia & "'"
    frmReferencia.Label1.Caption = "Lineas"
    frmReferencia.Show vbModal
    Adodc3.Close
    If vGUtil(1) <> "" Then TxArt1 = (vGUtil(1))
ElseIf Option1(1).Value Then
    Adodc3.Open "SELECT GRU_CODIGO,GRU_NOMBRE  FROM GRUPO WHERE FAM_CODIGO = '" & TxFamilia & "' AND LIN_CODIGO = '" & TxLinea & "'", cconexcom, adOpenStatic, adLockOptimistic
    frmReferencia.Conectar Adodc3, "SELECT GRU_CODIGO,GRU_NOMBRE  FROM GRUPO WHERE FAM_CODIGO = '" & TxFamilia & "' AND LIN_CODIGO = '" & TxLinea & "'"
    frmReferencia.Label1.Caption = "Grupo"
    frmReferencia.Show vbModal
    Adodc3.Close
    If vGUtil(1) <> "" Then TxArt1 = (vGUtil(1))
End If
End Sub

Private Sub TxArt1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then TxArt1_DblClick
End Sub

Private Sub TxArt2_DblClick()
Dim Adodc3 As ADODB.Recordset
Set Adodc3 = New ADODB.Recordset
If Option1(0).Value Then
    Adodc3.Open "SELECT Acodigo,adescri  FROM MAeArt", cconexcom, adOpenStatic, adLockOptimistic
    frmReferencia.Conectar Adodc3, "SELECT Acodigo,adescri  FROM MAeArt"
    frmReferencia.Label1.Caption = "Articulos"
    frmReferencia.Show vbModal
    Adodc3.Close
    If vGUtil(1) <> "" Then TxArt2 = (vGUtil(1))
ElseIf Option1(3).Value Then
    Adodc3.Open "SELECT FAM_CODIGO,FAM_NOMBRE  FROM FAMILIA", cconexcom, adOpenStatic, adLockOptimistic
    frmReferencia.Conectar Adodc3, "SELECT FAM_CODIGO,FAM_NOMBRE  FROM FAMILIA"
    frmReferencia.Label1.Caption = "Familias"
    frmReferencia.Show vbModal
    Adodc3.Close
    If vGUtil(1) <> "" Then TxArt2 = (vGUtil(1))
ElseIf Option1(2).Value Then
    Adodc3.Open "SELECT LIN_CODIGO,LIN_NOMBRE  FROM LINEAS WHERE FAM_CODIGO = '" & TxFamilia & "'", cconexcom, adOpenStatic, adLockOptimistic
    frmReferencia.Conectar Adodc3, "SELECT LIN_CODIGO,LIN_NOMBRE  FROM LINEAS WHERE FAM_CODIGO = '" & TxFamilia & "'"
    frmReferencia.Label1.Caption = "Lineas"
    frmReferencia.Show vbModal
    Adodc3.Close
    If vGUtil(1) <> "" Then TxArt2 = (vGUtil(1))
ElseIf Option1(1).Value Then
    Adodc3.Open "SELECT GRU_CODIGO,GRU_NOMBRE  FROM GRUPO WHERE FAM_CODIGO = '" & TxFamilia & "' AND LIN_CODIGO = '" & TxLinea & "'", cconexcom, adOpenStatic, adLockOptimistic
    frmReferencia.Conectar Adodc3, "SELECT GRU_CODIGO,GRU_NOMBRE  FROM GRUPO WHERE FAM_CODIGO = '" & TxFamilia & "' AND LIN_CODIGO = '" & TxLinea & "'"
    frmReferencia.Label1.Caption = "Grupo"
    frmReferencia.Show vbModal
    Adodc3.Close
    If vGUtil(1) <> "" Then TxArt2 = (vGUtil(1))
End If
End Sub

Private Sub TxArt2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then TxArt2_DblClick
End Sub

Private Sub Rep_Opc(cCodIni As String, cCodFin As String, cAlma As String, dFecIni As Date, cFecFin As Date)
Dim Rep1 As ADODB.Recordset
Dim Rep2 As ADODB.Recordset
Dim Rep3 As ADODB.Recordset
Dim cS As String, WGRUPO As String, WFAMILIA As String
Dim WMODELO As String, WTIPO As String

Set Rep1 = New ADODB.Recordset
Set Rep2 = New ADODB.Recordset
Set Rep3 = New ADODB.Recordset

cConexAux.Execute "Delete From Vent_Cli"

cS = "Select * From FacDet A Inner Join FacCab B on "
cS = cS & "  A.DFTD = B.CFTD AND A.DFNUMSER = B.CFNUMSER AND A.DFNUMDOC = B.CFNUMDOC"
cS = cS & " Where CFESTADO <> 'A' and DFTR <> 'T' and DFALMA >= '" & cAlma & "'  and "
If Option1(0).Value Then 'CFVENDE  'CFCODCLI
    cS = cS & "DFCODIGO >= '" & cCodIni & "' and DFCODIGO <= '" & cCodFin & "' and "
Else
End If
cS = cS & "CFFECDOC >= #" & Format(dFecIni, "mm/dd/yyyy") & "# and CFFECDOC <= #" & Format(cFecFin, "mm/dd/yyyy") & "# AND CFPUNVEN >= '" & CmbPV1.text & "' AND CFPUNVEN <= '" & CmbPV1.text & "'"

Rep1.Open cS, cconexcom, adOpenStatic
 
Rep2.Open "SELECT * FROM MAEART", cconexcom, adOpenStatic

Rep3.Open "SELECT * FROM VENT_CLI", cConexAux, adOpenDynamic, adLockOptimistic

Do While Not Rep1.EOF
    WGRUPO = "": WFAMILIA = "": WMODELO = "": WTIPO = ""
    If Not Rep2.EOF Then
        Rep2.MoveFirst
        Rep2.Filter = "ACODIGO = '" & Rep1("DFCODIGO") & "'"
    End If
    If Not Rep2.EOF Then
        If Not IsNull(Rep2("AGRUPO")) Then WGRUPO = Rep2("AGRUPO")
        If Not IsNull(Rep2("AFAMILIA")) Then WFAMILIA = Rep2("AFAMILIA")
        If Not IsNull(Rep2("AMODELO")) Then WMODELO = Rep2("AMODELO")
        If Not IsNull(Rep2("ATIPO")) Then WTIPO = Rep2("ATIPO")
    End If
    Rep2.Filter = ""
    Rep3.AddNew
    Rep3("VENCODAGE") = Rep1("CFPUNVEN")
    Rep3("VENTD") = Rep1("DFTD")
    Rep3("VENSERDOC") = Rep1("DFNUMSER") & Rep1("DFNUMDOC")
    Rep3("VENFECDOC") = Rep1("CFFECDOC")
    Rep3("VENCODIGO") = Rep1("DFCODIGO")
    Rep3("VENSERIE") = IIf(Not IsNull(Rep1("DFSERIE")) And Trim(Rep1("DFSERIE")) <> "", Rep1("DFSERIE"), " ")
    Rep3("VENVENDE") = Rep1("CFVENDE")
    Rep3("VENCODVEN") = Mid(Dev_RegVal("TABAYU", "TCOD = '27' AND TCLAVE = '" & Rep1("CFVENDE") & "'", "TDESCRI", cconexcom), 1, 25)
    Rep3("VENGRUPO") = IIf(Trim(WGRUPO) = "", "  ", WGRUPO)
    Rep3("VENFAMILIA") = IIf(Trim(WFAMILIA) = "", "  ", WFAMILIA)
    Rep3("VENMODELO") = IIf(Trim(WMODELO) = "", "  ", WMODELO)
    Rep3("VENCODCLI") = Rep1("CFCODCLI")
    Rep3("VENMONEDA") = Rep1("CFCODMON")
    Rep3("VENCANTID") = Rep1("DFCANTID")
    If Combo2.ListIndex = 0 Then
         Rep3("VENIGV") = IIf(Rep1("CFCODMON") = "MN", Rep1("DFIGV"), Round(Rep1("DFIGV") * Rep1("CFTIPCAM"), 2))
    Else
         Rep3("VENIGV") = IIf(Rep1("CFCODMON") = "MN", Round(Rep1("DFIGV") / Rep1("CFTIPCAM"), 2), Rep1("DFIGV"))
    End If
    Rep3("VENIMPUS") = Rep1("DFIMPUS")
    Rep3("VENIMPMN") = Rep1("DFIMPMN")
    Rep3("VENTIPO") = IIf(Trim(WTIPO) = "", "  ", WTIPO)
    Rep3("VENTIPCAM") = Rep1("CFTIPCAM")
    Rep3.UpdateBatch
    Rep3.Requery
    Rep1.MoveNext
    If Rep1.EOF Then Exit Do
Loop
Rep3.Close: Rep1.Close: Rep2.Close
End Sub

Private Sub Limp_Tex()
TxArt1 = "": TxArt2 = ""
End Sub

Private Sub TxFamilia_DblClick()
Dim Adodc3 As ADODB.Recordset
Set Adodc3 = New ADODB.Recordset
Adodc3.Open "SELECT FAM_CODIGO,FAM_NOMBRE  FROM FAMILIA", cconexcom, adOpenStatic, adLockOptimistic
frmReferencia.Conectar Adodc3, "SELECT FAM_CODIGO,FAM_NOMBRE  FROM FAMILIA"
frmReferencia.Label1.Caption = "Familias"
frmReferencia.Show vbModal
Adodc3.Close
If vGUtil(1) <> "" Then TxFamilia = (vGUtil(1))
End Sub

Private Sub TxFamilia_GotFocus()
Enfoque TxFamilia
End Sub

Private Sub TxFamilia_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then TxFamilia_DblClick
End Sub

Private Sub TxFamilia_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If existe(1, TxFamilia, "familia", "fam_codigo", False) Then
         SendKeys "{tab}"
    Else
        MsgBox "La Familia no existe", vbInformation, "Información"
        TxFamilia.SetFocus
    End If
End If
End Sub

Private Sub TxLinea_DblClick()
Dim Adodc3 As ADODB.Recordset
Set Adodc3 = New ADODB.Recordset
Adodc3.Open "SELECT LIN_CODIGO,LIN_NOMBRE  FROM LINEAS WHERE FAM_CODIGO = '" & TxFamilia & "'", cconexcom, adOpenStatic, adLockOptimistic
frmReferencia.Conectar Adodc3, "SELECT LIN_CODIGO,LIN_NOMBRE  FROM LINEAS WHERE FAM_CODIGO = '" & TxFamilia & "' "
frmReferencia.Label1.Caption = "Lineas"
frmReferencia.Show vbModal
Adodc3.Close
If vGUtil(1) <> "" Then TxLinea = (vGUtil(1))
End Sub

Private Sub TxLinea_GotFocus()
Enfoque TxLinea
End Sub

Private Sub TxLinea_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then TxLinea_DblClick
End Sub

Private Sub TxLinea_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If existe(1, TxLinea, "Lineas", "lin_codigo", False, TxFamilia, "fam_codigo") Then
         SendKeys "{tab}"
    Else
        MsgBox "La Linea no existe", vbInformation, "Información"
        TxLinea.SetFocus
    End If
End If
End Sub
