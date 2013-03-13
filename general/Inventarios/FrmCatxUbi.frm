VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form FrmCatxUbi 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Catálogo por Ubicación"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5970
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   5970
   StartUpPosition =   3  'Windows Default
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   3735
      Top             =   2340
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton CmdS 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   675
      Left            =   1224
      Picture         =   "FrmCatxUbi.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2088
      Width           =   775
   End
   Begin VB.CommandButton CmdA 
      Caption         =   "&Aceptar"
      Height          =   675
      Left            =   252
      Picture         =   "FrmCatxUbi.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2088
      Width           =   775
   End
   Begin VB.Frame Frame1 
      Height          =   2772
      Left            =   108
      TabIndex        =   4
      Top             =   144
      Width           =   5784
      Begin VB.CheckBox Check1 
         Caption         =   "Listar Todos"
         Height          =   264
         Left            =   180
         TabIndex        =   7
         Top             =   1476
         Width           =   1920
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   1
         Left            =   2130
         MaxLength       =   12
         TabIndex        =   1
         Top             =   870
         Width           =   1116
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   0
         Left            =   2145
         MaxLength       =   12
         TabIndex        =   0
         Top             =   420
         Width           =   1116
      End
      Begin VB.Image Image1 
         Height          =   2292
         Left            =   3744
         Picture         =   "FrmCatxUbi.frx":0884
         Stretch         =   -1  'True
         Top             =   252
         Width           =   1824
      End
      Begin VB.Label Label2 
         Caption         =   "Al Cod. de Ubicación"
         Height          =   300
         Left            =   195
         TabIndex        =   6
         Top             =   900
         Width           =   1785
      End
      Begin VB.Label Label1 
         Caption         =   "Del Cod. de Ubicación"
         Height          =   225
         Left            =   180
         TabIndex        =   5
         Top             =   465
         Width           =   1785
      End
   End
End
Attribute VB_Name = "FrmCatxUbi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Adodc3 As ADODB.Recordset



Private Sub CmdA_Click()
Dim cadena As String

If Check1.Value = 1 Then
   cadena = "{STKART.STALMA} = '" & VGAlma & "' AND ({MAEART.ADESCRI} <>'')"
Else
   If Text1(0).text = "" Or Text1(1).text = "" Then Exit Sub
   cadena = "{tabcasillero.tcasillero}>= '" & Text1(0).text & "' and {tabcasillero.tcasillero}<= '" & Text1(1).text & "' and {STKART.STALMA} = '" & VGAlma & "'"
End If


On Error GoTo error
    CrystalReport1.WindowTitle = "Inv145 -- Control de Inventarios"
    CrystalReport1.ReportFileName = cRutP & "inv145.rpt"
    Ubi_Tab CrystalReport1
    CrystalReport1.WindowShowPrintBtn = True
    CrystalReport1.WindowShowRefreshBtn = True
    CrystalReport1.WindowShowSearchBtn = True
    CrystalReport1.WindowShowPrintSetupBtn = True
    CrystalReport1.DiscardSavedData = True
    CrystalReport1.Destination = crptToWindow
    CrystalReport1.ReplaceSelectionFormula (cadena)
    CrystalReport1.Formulas(0) = "emp = '" & VGNemp & "'"
    CrystalReport1.Formulas(1) = "alma ='" & VGNomAlm & "'"
    If CrystalReport1.Status <> 2 Then CrystalReport1.Action = 1
    Exit Sub
error:
    MsgBox Err.Description

End Sub

Private Sub CmdS_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  Text1(0) = ""
  Text1(1) = ""
  central Me
End Sub

Private Sub Text1_DblClick(Index As Integer)
    Set Adodc3 = New ADODB.Recordset
    Adodc3.Open "SELECT DISTINCT TCASILLERO,TCODALM FROM TABCASILLERO WHERE TCODALM='" & VGAlma & "'", cConexCom, adOpenStatic, adLockOptimistic
    frmReferencia.Conectar Adodc3, "SELECT DISTINCT TCASILLERO,TCODALM FROM TABCASILLERO WHERE TCODALM='" & VGAlma & "'"
    frmReferencia.Label1.Caption = "Ubicacion de Articulos"
    frmReferencia.Show vbModal
    Adodc3.Close
    If vGUtil(1) <> "" Then
        Text1(Index) = vGUtil(1)
    End If
    If Text1(Index) <> "" Then Call Text1_KeyPress(Index, 13)
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys "{tab}"
   End If

End Sub
