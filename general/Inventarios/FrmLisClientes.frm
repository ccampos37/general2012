VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form FrmLisClientes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Clientes"
   ClientHeight    =   4470
   ClientLeft      =   3990
   ClientTop       =   2220
   ClientWidth     =   3825
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   3825
   Begin VB.Frame Frame1 
      Caption         =   "Parámetros"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   72
      TabIndex        =   0
      Top             =   0
      Width           =   3684
      Begin VB.OptionButton Option1 
         Caption         =   "Por R.U.C."
         Height          =   375
         Index           =   2
         Left            =   480
         TabIndex        =   3
         Top             =   1080
         Width           =   2055
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Por Razón Social"
         Height          =   375
         Index           =   1
         Left            =   480
         TabIndex        =   2
         Top             =   720
         Width           =   2055
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Por Código "
         Height          =   375
         Index           =   0
         Left            =   480
         TabIndex        =   1
         Top             =   360
         Width           =   2055
      End
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   2880
         Top             =   3000
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1095
      Left            =   96
      TabIndex        =   16
      Top             =   3312
      Width           =   3612
      Begin VB.CommandButton CmdS 
         Caption         =   "&Salir"
         Height          =   675
         Left            =   1800
         Picture         =   "FrmLisClientes.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   240
         Width           =   775
      End
      Begin VB.CommandButton CmdA 
         Caption         =   "&Aceptar"
         Height          =   675
         Left            =   600
         Picture         =   "FrmLisClientes.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   240
         Width           =   775
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Criterio"
      Height          =   1548
      Left            =   60
      TabIndex        =   4
      Top             =   1596
      Visible         =   0   'False
      Width           =   3684
      Begin VB.CheckBox Check2 
         Caption         =   "Listar Todo"
         Height          =   195
         Left            =   240
         TabIndex        =   15
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox TxA2 
         Height          =   285
         Left            =   1560
         TabIndex        =   8
         Text            =   "Text2"
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox TxA1 
         Height          =   285
         Left            =   1560
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Código Final   :"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Código Inicial  :"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   600
         Width           =   1455
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Criterio"
      Height          =   1488
      Left            =   72
      TabIndex        =   9
      Top             =   1656
      Visible         =   0   'False
      Width           =   3648
      Begin VB.CheckBox Check1 
         Caption         =   "Listar Todo"
         Height          =   195
         Left            =   240
         TabIndex        =   14
         Top             =   348
         Width           =   1935
      End
      Begin VB.TextBox TxC2 
         Height          =   285
         Left            =   1560
         TabIndex        =   11
         Text            =   "Text2"
         Top             =   1068
         Width           =   1335
      End
      Begin VB.TextBox TxC1 
         Height          =   285
         Left            =   1560
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   708
         Width           =   1335
      End
      Begin VB.Label Label12 
         Caption         =   "RUC  Final   :"
         Height          =   252
         Left            =   240
         TabIndex        =   13
         Top             =   1068
         Width           =   1332
      End
      Begin VB.Label Label11 
         Caption         =   "RUC  Inicial  :"
         Height          =   252
         Left            =   240
         TabIndex        =   12
         Top             =   708
         Width           =   1452
      End
   End
End
Attribute VB_Name = "FrmLisClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nSw As Integer, nSw2 As Integer
Dim CTIME As String
Dim nSw11 As Integer, nSw22 As Integer
Dim nSw111 As Integer, nSw222 As Integer
Dim nOpc As Integer, nOpcP As Integer
Dim cT As String

Private Sub Check1_Click()
If Check1.Value = 1 Then
    TxC1.Enabled = False
    TxC2.Enabled = False
Else
    TxC1.Enabled = True
    TxC2.Enabled = True
End If
End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
    TxA1.Enabled = False
    TxA2.Enabled = False
Else
    TxA1.Enabled = True
    TxA2.Enabled = True
End If
End Sub

Private Sub CmdA_Click()
If Option1(0).Value = True Then  ' Cliente
    If Check2.Value = 1 Then
        nOpc = 0
        ImpT
        Exit Sub
    End If
    
    If TxA1 <> "" Then
    If CodigoC(TxA1) Then
        MsgBox "Código no Existe", vbInformation, "Mensaje"
        TxA1.SetFocus
        Exit Sub
    End If
    End If
    
    If TxA2 <> "" Then
    If CodigoC(TxA2) Then
        MsgBox "Código no Existe", vbInformation, "Mensaje"
        TxA2.SetFocus
        Exit Sub
    End If
    End If
    
    If Trim(TxA1) > Trim(TxA2) Then
        MsgBox "El Código Inicial es Mayor al Código Final"
        TxA1.SetFocus
        Exit Sub
    Else
        nOpc = 0
        ImpT
    End If
ElseIf Option1(1).Value = True Then 'Razon Social
    nOpc = 1
    ImpT
ElseIf Option1(2).Value = True Then 'Ruc
    If Check1.Value = 1 Then
        nOpc = 2
        ImpT
        Exit Sub
    End If
    
    If TxC1 <> "" Then
    If Existe(1, TxC1, "MAECLI", "CNUMRUC", False) = False Then
        MsgBox "Código de RUC no Existe", vbInformation, "Mensaje"
        TxC1.SetFocus
        Exit Sub
    End If
    End If
    
    If TxC2 <> "" Then
    If Existe(1, TxC2, "MAECLI", "CNUMRUC", False) = False Then
        MsgBox "Código de RUC no Existe", vbInformation, "Mensaje"
        TxC2.SetFocus
        Exit Sub
    End If
    End If
    
    If Trim(TxC1) > Trim(TxC2) Then
        MsgBox "El RUC es Mayor al RUC Final"
        TxC1.SetFocus
        Exit Sub
    Else
        nOpc = 2
        ImpT
    End If
End If
End Sub

Private Sub ImpT()
On Error GoTo Err
Screen.MousePointer = 1
If nOpc = 0 Then
    CrystalReport1.WindowTitle = "Inv010 -- Control de Inventarios"
    CrystalReport1.ReportFileName = cRutP & "inv010.rpt"
    Call Ubi_Tab(CrystalReport1)
    If Check2.Value = 1 Then
        CrystalReport1.SelectionFormula = ""
    Else
        CrystalReport1.SelectionFormula = "{MAECLI.CCODCLI} >=  '" & Trim(TxA1) & "' and {MAECLI.CCODCLI} <=  '" & Trim(TxA2) & "'"
    End If
    CrystalReport1.WindowShowPrintBtn = True
    CrystalReport1.WindowShowRefreshBtn = True
    CrystalReport1.WindowShowSearchBtn = True
    CrystalReport1.WindowShowPrintSetupBtn = True
    CrystalReport1.WindowTop = 100
    CrystalReport1.WindowLeft = 150
    CrystalReport1.DiscardSavedData = True
    CrystalReport1.Destination = crptToWindow
    CrystalReport1.Action = 1
ElseIf nOpc = 1 Then
    CrystalReport1.WindowTitle = "Inv016 -- Sistema de Inventarios "
    CrystalReport1.ReportFileName = cRutP & "inv016.Rpt"
    Call Ubi_Tab(CrystalReport1)
    CrystalReport1.WindowShowPrintBtn = True
    CrystalReport1.WindowShowRefreshBtn = True
    CrystalReport1.WindowShowSearchBtn = True
    CrystalReport1.WindowShowPrintSetupBtn = True
    CrystalReport1.SelectionFormula = ""
    CrystalReport1.WindowTop = 100
    CrystalReport1.WindowLeft = 150
    CrystalReport1.DiscardSavedData = True
    CrystalReport1.Destination = crptToWindow
    CrystalReport1.Action = 1
ElseIf nOpc = 2 Then
    CrystalReport1.WindowTitle = "Inv099 -- Sistema de Inventarios "
    CrystalReport1.ReportFileName = cRutP & "inv099.rpt"                               'RucCli.Rpt"
    Call Ubi_Tab(CrystalReport1)
    If Check1.Value = 1 Then
        CrystalReport1.SelectionFormula = ""
    Else
        CrystalReport1.SelectionFormula = "{MAECLI.CNUMRUC} >=  '" & Trim(TxC1) & "' and {MAECLI.CNUMRUC} <=  '" & Trim(TxC2) & "'"
    End If
    CrystalReport1.WindowShowPrintBtn = True
    CrystalReport1.WindowShowRefreshBtn = True
    CrystalReport1.WindowShowSearchBtn = True
    CrystalReport1.WindowShowPrintSetupBtn = True
    CrystalReport1.WindowTop = 100
    CrystalReport1.WindowLeft = 150
    CrystalReport1.DiscardSavedData = True
    CrystalReport1.Destination = crptToWindow
    CrystalReport1.Action = 1
End If
  Screen.MousePointer = 1
  Exit Sub
Err:
    MsgBox "No se encontro el reporte", vbInformation, "Aviso"
    Screen.MousePointer = 1
End Sub

Private Sub CmdS_Click()
Unload Me
End Sub


Private Sub Form_Load()
central Me
nSw = 0
nSw2 = 0
nSw11 = 0
nSw22 = 0
nOpc = 0
nOpcP = 0
Limpiar
CTIME = Format(Time, "hh:mm:ss")
CrystalReport1.WindowTitle = "Sistema de Inventarios"
CrystalReport1.Formulas(0) = "Hora = '" & CTIME & "'"
CrystalReport1.Formulas(1) = "Empresa = '" & Mid(VGNemp, 1, 20) & "'"
End Sub

Private Sub Option1_Click(Index As Integer)
Select Case Index
Case 0 ' Articulo
    Frame2.Visible = True
    Frame7.Visible = True
    Frame7.Enabled = False
Case 1 ' Descripcion
    Frame2.Visible = False
    Frame7.Visible = True
    Frame7.Enabled = False
Case 2 'ruc
    Frame2.Visible = False
    Frame7.Visible = True
    Frame7.Enabled = True
End Select
Limpiar
End Sub

Private Sub TxA1_DblClick()
Static Adodc2 As ADODB.Recordset
Set Adodc2 = New ADODB.Recordset
    Select Case Index
     Case 0:

         Adodc2.Open "SELECT CCODCLI,CNOMCLI FROM MAECLI", cConexCom, adOpenStatic, adLockOptimistic
         frmReferencia.Conectar Adodc2, "SELECT CCODCLI,CNOMCLI FROM MAECLI"
         frmReferencia.Label1.Caption = "Maestro de Clientes"
         frmReferencia.Show vbModal
         Adodc2.Close
         If vGUtil(1) <> "" Then
           TxA1.text = (vGUtil(1))
           'Text1(1) = VGUTIL(2)
         End If
   End Select

End Sub

Private Sub TxA1_GotFocus()
Enfoque TxA1
End Sub

Private Sub TxA1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then TxA1_DblClick
End Sub

Private Sub TxA1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   
    If Trim(TxA1) <> "" Then
          If Existe(1, TxA1, "MAECLI", "CCODCLI", False) = False Then
            MsgBox "Codigo de Cliente no existe", vbInformation, mensaje1
            TxA1.SetFocus: Exit Sub
          End If
    End If
    TxA2.SetFocus: Exit Sub
    
Else
    If KeyAscii = 39 Or KeyAscii = 34 Then KeyAscii = 0
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End If
End Sub

Private Sub TxA2_DblClick()
Static Adodc2 As ADODB.Recordset
Set Adodc2 = New ADODB.Recordset
    Select Case Index
     Case 0:

         Adodc2.Open "SELECT CCODCLI,CNOMCLI FROM MAECLI", cConexCom, adOpenStatic, adLockOptimistic
         frmReferencia.Conectar Adodc2, "SELECT CCODCLI,CNOMCLI FROM MAECLI"
         frmReferencia.Label1.Caption = "Maestro de Clientes"
         frmReferencia.Show vbModal
         Adodc2.Close
         If vGUtil(1) <> "" Then
           TxA2.text = (vGUtil(1))
           'Text1(1) = VGUTIL(2)
         End If
   End Select
End Sub

Private Sub TxA2_GotFocus()
Enfoque TxA2
End Sub

Private Sub TxA2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then TxA2_DblClick
End Sub

Private Sub TxA2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Trim(TxA2) <> "" Then
          If Existe(1, TxA2, "MAECLI", "CCODCLI", False) = False Then
            MsgBox "Codigo de Cliente no existe", vbInformation, mensaje1
            TxA2.SetFocus: Exit Sub
          End If
    End If
    CmdA.SetFocus: Exit Sub
    
Else
    If KeyAscii = 39 Or KeyAscii = 34 Then KeyAscii = 0
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End If
End Sub

Private Sub TxC1_DblClick()
Static Adodc2 As ADODB.Recordset
Set Adodc2 = New ADODB.Recordset
    Select Case Index
     Case 0:
        Adodc2.Open "SELECT CNUMRUC,CNOMCLI FROM MAECLI", cConexCom, adOpenStatic, adLockOptimistic
        frmReferencia.Conectar Adodc2, "SELECT CNUMRUC,CNOMCLI FROM MAECLI"
        frmReferencia.Label1.Caption = "RUC de Clientes"
        frmReferencia.Show vbModal
        Adodc2.Close
         If vGUtil(1) <> "" Then
           TxC1.text = (vGUtil(1))
           'Text1(1) = VGUTIL(2)
         End If
   End Select
End Sub

Private Sub TxC1_GotFocus()
Enfoque TxC1
End Sub

Private Sub TxC1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then TxC1_DblClick
End Sub

Private Sub TxC1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Trim(TxC1) <> "" Then
          If Existe(1, TxC1, "MAECLI", "CNUMRUC", False) = False Then
            MsgBox "RUC de Cliente no existe", vbInformation, mensaje1
            TxC1.SetFocus: Exit Sub
          End If
    End If
    TxC2.SetFocus: Exit Sub
    
Else
    If KeyAscii = 39 Or KeyAscii = 34 Then KeyAscii = 0
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End If
End Sub

Private Sub TxC2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then TxC2_DblClick
End Sub
Private Sub TxC2_DblClick()
Static Adodc2 As ADODB.Recordset
Set Adodc2 = New ADODB.Recordset
    Select Case Index
     Case 0:
        Adodc2.Open "SELECT CNUMRUC,CNOMCLI FROM MAECLI", cConexCom, adOpenStatic, adLockOptimistic
        frmReferencia.Conectar Adodc2, "SELECT CNUMRUC,CNOMCLI FROM MAECLI"
        frmReferencia.Label1.Caption = "RUC de Clientes"
        frmReferencia.Show vbModal
        Adodc2.Close
         If vGUtil(1) <> "" Then
           TxC2.text = (vGUtil(1))
           'Text1(1) = VGUTIL(2)
         End If
   End Select
End Sub

Private Sub TxC2_GotFocus()
Enfoque TxC2
End Sub
Private Sub Limpiar()
TxA1 = ""
TxA2 = ""
TxC1 = ""
TxC2 = ""
End Sub

Private Function CodigoR(cCod As String) As Boolean     'Ruc
Dim cSelC As ADODB.Recordset, csql As String
If Trim(cCod) = "" Then
    MsgBox "Falta Codigo", vbInformation, "Mensaje"
    CodigoR = False
    Exit Function
End If
csql = "Select cnumruc from Maecli where cnumruc = '" & cCod & "'"
Set cSelC = New ADODB.Recordset
cSelC.Open csql, cConexCom, adOpenStatic
If cSelC.RecordCount > 0 Then
    CodigoR = False: cSelC.Close
    Exit Function
End If
CodigoR = True: cSelC.Close
End Function

Private Sub TxC2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Trim(TxC2) <> "" Then
          If Existe(1, TxC2, "MAECLI", "CNUMRUC", False) = False Then
            MsgBox "RUC de Cliente no existe", vbInformation, mensaje1
            TxC2.SetFocus: Exit Sub
          End If
    End If
    CmdA.SetFocus: Exit Sub
    
Else
    If KeyAscii = 39 Or KeyAscii = 34 Then KeyAscii = 0
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End If
End Sub
