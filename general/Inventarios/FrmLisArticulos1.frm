VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form FrmLisArticulos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Articulos"
   ClientHeight    =   5415
   ClientLeft      =   3855
   ClientTop       =   1485
   ClientWidth     =   3735
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   3735
   Begin VB.Frame Frame9 
      Height          =   1095
      Left            =   240
      TabIndex        =   39
      Top             =   4200
      Width           =   3255
      Begin VB.CommandButton CmdA 
         Caption         =   "&Aceptar"
         Height          =   675
         Left            =   480
         Picture         =   "FrmLisArticulos.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   240
         Width           =   775
      End
      Begin VB.CommandButton CmdS 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         Height          =   675
         Left            =   1920
         Picture         =   "FrmLisArticulos.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   240
         Width           =   775
      End
   End
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
      Height          =   2775
      Left            =   240
      TabIndex        =   0
      Top             =   0
      Width           =   3255
      Begin VB.OptionButton Option1 
         Caption         =   "Por Tipo Articulo"
         Height          =   375
         Index           =   7
         Left            =   480
         TabIndex        =   22
         Top             =   2160
         Width           =   2055
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Por Cta. Contable"
         Height          =   375
         Index           =   6
         Left            =   480
         TabIndex        =   21
         Top             =   1800
         Width           =   2055
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Por Cod. Fabricante"
         Height          =   375
         Index           =   5
         Left            =   480
         TabIndex        =   20
         Top             =   1440
         Width           =   2055
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Por Familia"
         Height          =   375
         Index           =   3
         Left            =   480
         TabIndex        =   3
         Top             =   1080
         Width           =   2055
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Por Descripción"
         Height          =   375
         Index           =   1
         Left            =   480
         TabIndex        =   2
         Top             =   720
         Width           =   2055
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Por Código Interno"
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
         _Version        =   262150
      End
   End
   Begin VB.Frame Frame6 
      Height          =   1455
      Left            =   240
      TabIndex        =   19
      Top             =   2760
      Visible         =   0   'False
      Width           =   3255
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "FrmLisArticulos.frx":0884
         Left            =   1350
         List            =   "FrmLisArticulos.frx":0891
         Style           =   2  'Dropdown List
         TabIndex        =   42
         Top             =   600
         Width           =   1500
      End
      Begin VB.CheckBox Check7 
         Caption         =   "Listar Todo"
         Height          =   255
         Left            =   240
         TabIndex        =   38
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "Tipo        :"
         Height          =   255
         Left            =   360
         TabIndex        =   43
         Top             =   660
         Width           =   960
      End
   End
   Begin VB.Frame Frame5 
      Height          =   1455
      Left            =   240
      TabIndex        =   14
      Top             =   2760
      Visible         =   0   'False
      Width           =   3255
      Begin VB.CheckBox Check1 
         Caption         =   "Listar Todo"
         Height          =   255
         Left            =   240
         TabIndex        =   35
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox TxF1 
         Height          =   285
         Left            =   1560
         TabIndex        =   16
         Text            =   "Text1"
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox TxF2 
         Height          =   285
         Left            =   1560
         TabIndex        =   15
         Text            =   "Text2"
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "De la Familia  :"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "A la Familia    :"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   960
         Width           =   1095
      End
   End
   Begin VB.Frame Frame7 
      Height          =   1455
      Left            =   240
      TabIndex        =   23
      Top             =   2760
      Visible         =   0   'False
      Width           =   3255
      Begin VB.CheckBox Check6 
         Caption         =   "Listar Todo"
         Height          =   255
         Left            =   240
         TabIndex        =   37
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox TxC2 
         Height          =   285
         Left            =   1560
         TabIndex        =   25
         Text            =   "Text2"
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox TxC1 
         Height          =   285
         Left            =   1560
         TabIndex        =   24
         Text            =   "Text1"
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label12 
         Caption         =   "Cod. Fab. Final   :"
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label11 
         Caption         =   "Cod. Fab. Inicial  :"
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   600
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1455
      Left            =   240
      TabIndex        =   4
      Top             =   2760
      Visible         =   0   'False
      Width           =   3255
      Begin VB.CheckBox Check4 
         Caption         =   "Listar Todo"
         Height          =   255
         Left            =   240
         TabIndex        =   34
         Top             =   240
         Width           =   1455
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
         Caption         =   "Al Artículo    :"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Del Artículo  :"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   600
         Width           =   1455
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1455
      Left            =   240
      TabIndex        =   9
      Top             =   2760
      Visible         =   0   'False
      Width           =   3255
      Begin VB.CheckBox Check3 
         Caption         =   "Listar Todo"
         Height          =   255
         Left            =   240
         TabIndex        =   33
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox TxG1 
         Height          =   285
         Left            =   1560
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox TxG2 
         Height          =   285
         Left            =   1560
         TabIndex        =   10
         Text            =   "Text2"
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Del Grupo  :"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Al Grupo    :"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   960
         Width           =   1095
      End
   End
   Begin VB.Frame Frame8 
      Height          =   1455
      Left            =   240
      TabIndex        =   28
      Top             =   2760
      Visible         =   0   'False
      Width           =   3255
      Begin VB.CheckBox Check5 
         Caption         =   "Listar Todo"
         Height          =   255
         Left            =   240
         TabIndex        =   36
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox TxCon2 
         Height          =   285
         Left            =   1560
         TabIndex        =   30
         Text            =   "Text2"
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox TxCon1 
         Height          =   285
         Left            =   1560
         TabIndex        =   29
         Text            =   "Text1"
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label14 
         Caption         =   "Cta. Cont. Final  :"
         Height          =   255
         Left            =   240
         TabIndex        =   32
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label13 
         Caption         =   "Cta. Cont. Inicial :"
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   600
         Width           =   1455
      End
   End
End
Attribute VB_Name = "FrmLisArticulos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nSw As Integer, nSw2 As Integer
Dim nSw11 As Integer, nSw22 As Integer
Dim nSw111 As Integer, nSw222 As Integer
Dim nOpc As Integer, nOpcP As Integer
Dim CTIME As String
Dim cT As String

Private Sub Check1_Click()
If Check1.Value = 1 Then
    TxF1.Enabled = False
    TxF2.Enabled = False
Else
    TxF1.Enabled = True
    TxF2.Enabled = True
End If
End Sub

Private Sub Check4_Click()
If Check4.Value = 1 Then
    TxA1.Enabled = False
    TxA2.Enabled = False
Else
    TxA1.Enabled = True
    TxA2.Enabled = True
End If
End Sub

Private Sub Check5_Click()
If Check5.Value = 1 Then
    TxCon1.Enabled = False
    TxCon2.Enabled = False
Else
    TxCon1.Enabled = True
    TxCon2.Enabled = True
End If
End Sub

Private Sub Check6_Click()
If Check6.Value = 1 Then
    TxC1.Enabled = False
    TxC2.Enabled = False
Else
    TxC1.Enabled = True
    TxC2.Enabled = True
End If
End Sub

Private Sub Check7_Click()
If Check7.Value = 1 Then
    Combo1.Enabled = False
Else
    Combo1.Enabled = True
End If
End Sub

Private Sub CmdA_Click()
If Option1(0).Value = True Then  ' Articulo
    If Check4.Value = 1 Then
        nOpc = 0
        ImpT
        Exit Sub
    End If
    If codigo(TxA1) Then
        MsgBox "Código no Existe", vbInformation, "Mensaje"
        TxA1.SetFocus
        Exit Sub
    End If
    If codigo(TxA2) Then
        MsgBox "Código no Existe", vbInformation, "Mensaje"
        TxA2.SetFocus
        Exit Sub
    End If
    
    If Trim(TxA1) > Trim(TxA2) Then
        MsgBox "El Articulo Inicial es Mayor al Articulo Final"
        TxA1.SetFocus
        Exit Sub
    Else
        nOpc = 0
        ImpT
    End If
ElseIf Option1(1).Value = True Then 'Descripcion
    nOpc = 1
    ImpT
ElseIf Option1(3).Value = True Then 'Familia
    If Check1.Value = 1 Then
        nOpc = 3
        ImpT
        Exit Sub
    End If
    If TxF1 <> "" Then
    If fFam(TxF1) = "" Then
        MsgBox "Código no Existe", vbInformation, "Mensaje"
        TxF1.SetFocus
        Exit Sub
    End If
    End If
    If TxF2 <> "" Then
    If fFam(TxF2) = "" Then
        MsgBox "Código no Existe", vbInformation, "Mensaje"
        TxF2.SetFocus
        Exit Sub
    End If
    End If
    
    If Trim(TxF1) > Trim(TxF2) Then
        MsgBox "La Familia Inicial es Mayor a la Familia Final", vbExclamation, "Aviso"
        TxF1.SetFocus
        Exit Sub
    Else
        nOpc = 3
        ImpT
    End If
ElseIf Option1(5).Value = True Then 'Cod. Fab.
    If Check6.Value = 1 Then
        nOpc = 5
        ImpT
        Exit Sub
    End If
   If TxC1 <> "" Then
    If CODIGO2(TxC1) Then
        MsgBox "Código no Existe", vbInformation, "Mensaje"
        TxC1.SetFocus
        Exit Sub
    End If
    End If
    If TxC2 <> "" Then
    If CODIGO2(TxC2) Then
        MsgBox "Código no Existe", vbInformation, "Mensaje"
        TxC2.SetFocus
        Exit Sub
    End If
    End If
    
    If Trim(TxC1) > Trim(TxC2) Then
        MsgBox "El Cod. Fabricante Inicial es Mayor al Cod. Fabr. Final Final", vbExclamation, "Aviso"
        TxC1.SetFocus
        Exit Sub
    Else
        nOpc = 5
        ImpT
    End If
    
ElseIf Option1(6).Value = True Then 'Cta. Cont.
    If Check5.Value = 1 Then
        nOpc = 6
        ImpT
        Exit Sub
    End If
    If TxCon1 <> "" Then
    
    If Devolver_Dato(1, TxCon1, "CUENTA_CONTABLE", "COD_CUENTA", "DES_CUENTA") = "" Then
        MsgBox "Código no Existe", vbInformation, "Mensaje"
        TxCon1.SetFocus
        Exit Sub
    End If
    End If
    
    If TxCon2 <> "" Then
    If Devolver_Dato(1, TxCon2, "CUENTA_CONTABLE", "COD_CUENTA", "DES_CUENTA") = "" Then
        MsgBox "Código no Existe", vbInformation, "Mensaje"
        TxCon2.SetFocus
        Exit Sub
    End If
    End If
    
    If Trim(TxCon1) > Trim(TxCon2) Then
        MsgBox "La Cta. Cont. Inicial es Mayor al Cta.Cont. Final", vbExclamation, "Aviso"
        TxCon1.SetFocus
        Exit Sub
    Else
        nOpc = 6
        ImpT
    End If

ElseIf Option1(7).Value = True Then 'Tipo
        nOpc = 7
        ImpT
End If
End Sub

Private Sub ImpT()
If nOpc = 0 Then  'Codigo
    CrystalReport1.ReportFileName = cRutP & "\CodArti.Rpt"
    Call Ubi_Tab(CrystalReport1)
    If Check4.Value = 1 Then
        CrystalReport1.SelectionFormula = ""
    Else
        CrystalReport1.SelectionFormula = "{MAEART.ACODIGO} >=  '" & Trim(TxA1) & "' and {MAEART.ACODIGO} <=  '" & Trim(TxA2) & "'"
    End If
    CrystalReport1.WindowTop = 100
    CrystalReport1.WindowLeft = 150
    CrystalReport1.DiscardSavedData = True
    CrystalReport1.Destination = crptToWindow
    If CrystalReport1.Status <> 2 Then CrystalReport1.Action = 1
    
ElseIf nOpc = 1 Then  'Descripcion
    CrystalReport1.ReportFileName = cRutP & "\DesArti.Rpt"
    Call Ubi_Tab(CrystalReport1)
    CrystalReport1.SelectionFormula = ""
    CrystalReport1.WindowTop = 100
    CrystalReport1.WindowLeft = 150
    CrystalReport1.DiscardSavedData = True
    CrystalReport1.Destination = crptToWindow
    If CrystalReport1.Status <> 2 Then CrystalReport1.Action = 1
ElseIf nOpc = 3 Then
    CrystalReport1.ReportFileName = cRutP & "\CodXFam.Rpt"
    Call Ubi_Tab(CrystalReport1)
    If Check1.Value = 1 Then
        CrystalReport1.SelectionFormula = ""
    Else
        CrystalReport1.SelectionFormula = "{MAEART.AFAMILIA} >=  '" & Trim(TxF1) & "' and {MAEART.AFAMILIA} <=  '" & Trim(TxF2) & "'"
    End If
    CrystalReport1.WindowTop = 100
    CrystalReport1.WindowLeft = 150
    CrystalReport1.DiscardSavedData = True
    CrystalReport1.Destination = crptToWindow
    If CrystalReport1.Status <> 2 Then CrystalReport1.Action = 1
ElseIf nOpc = 5 Then
    CrystalReport1.ReportFileName = cRutP & "\CodXFabr.Rpt"
    Call Ubi_Tab(CrystalReport1)
    If Check6.Value = 1 Then
        CrystalReport1.SelectionFormula = ""
    Else
        CrystalReport1.SelectionFormula = "{MAEART.ACODIGO2} >=  '" & Trim(TxC1) & "' and {MAEART.ACODIGO2} <=  '" & Trim(TxC2) & "'"
    End If
    CrystalReport1.WindowTop = 100
    CrystalReport1.WindowLeft = 150
    CrystalReport1.DiscardSavedData = True
    CrystalReport1.Destination = crptToWindow
    If CrystalReport1.Status <> 2 Then CrystalReport1.Action = 1

ElseIf nOpc = 6 Then
    CrystalReport1.ReportFileName = cRutP & "\CodCta.Rpt"
    Call Ubi_Tab(CrystalReport1)
    If Check5.Value = 1 Then
        CrystalReport1.SelectionFormula = ""
    Else
        CrystalReport1.SelectionFormula = "{MAEART.ACUENTA} >=  '" & Trim(TxCon1) & "' and {MAEART.ACUENTA} <=  '" & Trim(TxCon2) & "'"
    End If
    CrystalReport1.WindowTop = 100
    CrystalReport1.WindowLeft = 150
    CrystalReport1.DiscardSavedData = True
    CrystalReport1.Destination = crptToWindow
    If CrystalReport1.Status <> 2 Then CrystalReport1.Action = 1
ElseIf nOpc = 7 Then
    CrystalReport1.ReportFileName = cRutP & "\CodXTipo.Rpt"
    Call Ubi_Tab(CrystalReport1)
    If Check7.Value = 1 Then
        CrystalReport1.SelectionFormula = ""
    Else
        If Combo1.ListIndex = 0 Then
        CrystalReport1.SelectionFormula = "{MAEART.ATIPO} =  'I'"
        ElseIf Combo1.ListIndex = 1 Then
        CrystalReport1.SelectionFormula = "{MAEART.ATIPO} =  'N'"
        ElseIf Combo1.ListIndex = 2 Then
        CrystalReport1.SelectionFormula = "{MAEART.ATIPO} =  'S'"
        End If
    End If
    CrystalReport1.WindowTop = 100
    CrystalReport1.WindowLeft = 150
    CrystalReport1.DiscardSavedData = True
    CrystalReport1.Destination = crptToWindow
    If CrystalReport1.Status <> 2 Then CrystalReport1.Action = 1
End If
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

Frame2.Visible = True
Frame3.Visible = False
'Frame4.Visible = False
Frame5.Visible = False
Frame6.Visible = False
Frame7.Visible = False
Frame8.Visible = False
Option1(0).Value = True

CTIME = Format(Time, "hh:mm:ss")
CrystalReport1.WindowTitle = "Sistema  de Ventas"
CrystalReport1.Formulas(0) = "Hora = '" & CTIME & "'"
CrystalReport1.Formulas(1) = "Empresa = '" & Mid(vGNomEmp, 1, 20) & "'"
End Sub

Private Sub Option1_Click(Index As Integer)
Select Case Index
Case 0 ' Articulo
    Frame2.Visible = True
    Frame3.Visible = False
    'Frame4.Visible = False
    Frame5.Visible = False
    Frame6.Visible = False
    Frame7.Visible = False
    Frame8.Visible = False
Case 1 ' Descripcion
    Frame2.Visible = False
    Frame3.Visible = False
    'Frame4.Visible = False
    Frame5.Visible = False
    Frame6.Visible = False
    Frame7.Visible = False
    Frame8.Visible = False

Case 2 'grupo
    Frame3.Visible = True
    Frame2.Visible = False
    'Frame4.Visible = False
    Frame5.Visible = False
    Frame6.Visible = False
    Frame7.Visible = False
    Frame8.Visible = False
Case 3 'familia
    'Frame4.Visible = False
    Frame3.Visible = False
    Frame2.Visible = False
    Frame5.Visible = True
    Frame6.Visible = False
    Frame7.Visible = False
    Frame8.Visible = False
Case 4 'linea
    Frame5.Visible = False
    Frame3.Visible = False
    'Frame4.Visible = True
    Frame2.Visible = False
    Frame6.Visible = False
    Frame7.Visible = False
    Frame8.Visible = False
Case 5 ' Cod. Fab.
    Frame6.Visible = False
    Frame3.Visible = False
    'Frame4.Visible = False
    Frame5.Visible = False
    Frame2.Visible = False
    Frame7.Visible = True
    Frame8.Visible = False
Case 6 ' Cta Con
    Frame6.Visible = False
    Frame3.Visible = False
    'Frame4.Visible = False
    Frame5.Visible = False
    Frame2.Visible = False
    Frame7.Visible = False
    Frame8.Visible = True
Case 7 ' Tipo
    Frame6.Visible = True
    Frame3.Visible = False
    'Frame4.Visible = False
    Frame5.Visible = False
    Frame2.Visible = False
    Frame7.Visible = False
    Frame8.Visible = False
End Select
Limpiar
End Sub

Private Sub TxA1_DblClick()
Static Adodc2 As adodb.Recordset
Set Adodc2 = New adodb.Recordset

    Select Case Index
     Case 0:

         Adodc2.Open "SELECT ACODIGO,ADESCRI FROM MAEART", cConexCom, adOpenStatic, adLockOptimistic
         frmReferencia.Conectar Adodc2, "SELECT ACODIGO,ADESCRI FROM MAEART"
         frmReferencia.Label1.Caption = "Maestro de Articulos"
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
          If existe(1, TxA1, "MAEART", "ACODIGO", False) = False Then
            MsgBox "Codigo de Articulo no existe", vbInformation, mensaje1
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
Static Adodc2 As adodb.Recordset
Set Adodc2 = New adodb.Recordset

    Select Case Index
     Case 0:

         Adodc2.Open "SELECT ACODIGO,ADESCRI FROM MAEART", cConexCom, adOpenStatic, adLockOptimistic
         frmReferencia.Conectar Adodc2, "SELECT ACODIGO,ADESCRI FROM MAEART"
         frmReferencia.Label1.Caption = " Articulos"
         frmReferencia.Show vbModal
         Adodc2.Close
         If vGUtil(1) <> "" Then
           TxA2.text = (vGUtil(1))
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
          If existe(1, TxA2, "MAEART", "ACODIGO", False) = False Then
            MsgBox "Codigo de Articulo no existe", vbInformation, mensaje1
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
Static Adodc2 As adodb.Recordset
Set Adodc2 = New adodb.Recordset

    Select Case Index
     Case 0:
         Adodc2.Open "Select  DISTINCT ACODIGO2 from MAEART where ACODIGO2<>''", cConexCom, adOpenStatic, adLockOptimistic
         frmReferencia1.Conectar Adodc2, "Select  DISTINCT ACODIGO2 from MAEART where ACODIGO2<>  ''"
         frmReferencia1.Label1.Caption = "Codigo de Fabricantes"
         frmReferencia1.Show vbModal
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
          If existe(1, TxC1, "MAEART", "ACODIGO2", False) = False Then
             MsgBox "Codigo de Proveedor no existe", vbInformation, mensaje1
             TxC1.SetFocus: Exit Sub
          End If
    End If
    TxC2.SetFocus: Exit Sub
    
Else
    If KeyAscii = 39 Or KeyAscii = 34 Then KeyAscii = 0
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End If
End Sub

Private Sub TxC2_DblClick()
Static Adodc2 As adodb.Recordset
Set Adodc2 = New adodb.Recordset

    Select Case Index
     Case 0:
         Adodc2.Open "Select  DISTINCT ACODIGO2 from MAEART where ACODIGO2<>''", cConexCom, adOpenStatic, adLockOptimistic
         frmReferencia1.Conectar Adodc2, "Select  DISTINCT ACODIGO2 from MAEART where ACODIGO2<>''"
         frmReferencia1.Label1.Caption = "Codigo de Proveedores"
         frmReferencia1.Show vbModal
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

Private Sub TxC2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then TxC2_DblClick
End Sub

Private Sub TxC2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   
    If Trim(TxC2) <> "" Then
          If existe(1, TxC2, "MAEART", "ACODIGO2", False) = False Then
            MsgBox "Codigo de Proveedor no existe", vbInformation, mensaje1
            TxC2.SetFocus: Exit Sub
          End If
    End If
    CmdA.SetFocus: Exit Sub
    
Else
    If KeyAscii = 39 Or KeyAscii = 34 Then KeyAscii = 0
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End If

End Sub

Private Sub TxCon1_DblClick()
Static Adodc2 As adodb.Recordset
Set Adodc2 = New adodb.Recordset

    Select Case Index
     Case 0:
         Adodc2.Open "Select COD_CUENTA,DES_CUENTA FROM CUENTA_CONTABLE", cConexCom, adOpenStatic, adLockOptimistic
         frmReferencia.Conectar Adodc2, "Select COD_CUENTA,DES_CUENTA FROM CUENTA_CONTABLE"
         frmReferencia.Label1.Caption = "Cuenta Contable"
         frmReferencia.Show vbModal
         Adodc2.Close
         If vGUtil(1) <> "" Then
           TxCon1.text = (vGUtil(1))
           'Text1(1) = VGUTIL(2)
         End If
             
   End Select

End Sub

Private Sub TxCon1_GotFocus()
Enfoque TxCon1
End Sub

Private Sub TxCon1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then TxCon1_DblClick
End Sub

Private Sub TxCon1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Trim(TxCon1) <> "" Then
          If existe(1, TxCon1, "CUENTA_CONTABLE", "COD_CUENTA", False) = False Then
            MsgBox "Codigo de Cuenta no existe", vbInformation, mensaje1
            TxCon1.SetFocus: Exit Sub
          End If
    End If
    TxCon2.SetFocus: Exit Sub
    
Else
    If KeyAscii = 39 Or KeyAscii = 34 Then KeyAscii = 0
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End If
End Sub

Private Sub TxCon2_DblClick()
Static Adodc2 As adodb.Recordset
Set Adodc2 = New adodb.Recordset

    Select Case Index
     Case 0:
         Adodc2.Open "Select COD_CUENTA,DES_CUENTA FROM CUENTA_CONTABLE", cConexCom, adOpenStatic, adLockOptimistic
         frmReferencia.Conectar Adodc2, "Select COD_CUENTA,DES_CUENTA FROM CUENTA_CONTABLE"
         frmReferencia.Label1.Caption = "Cuenta Contable"
         frmReferencia.Show vbModal
         Adodc2.Close
         If vGUtil(1) <> "" Then
           TxCon2.text = (vGUtil(1))
         End If
             
   End Select
End Sub

Private Sub TxCon2_GotFocus()
Enfoque TxCon2
End Sub

Private Sub TxCon2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then TxCon2_DblClick
End Sub

Private Sub TxCon2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Trim(TxCon2) <> "" Then
          If existe(1, TxCon2, "CUENTA_CONTABLE", "COD_CUENTA", False) = False Then
            MsgBox "Codigo de Cuenta no existe", vbInformation, mensaje1
            TxCon2.SetFocus: Exit Sub
          End If
    End If
    CmdA.SetFocus: Exit Sub
    
Else
    If KeyAscii = 39 Or KeyAscii = 34 Then KeyAscii = 0
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End If
End Sub

Private Sub TxF1_DblClick()
Static Adodc2 As adodb.Recordset
Set Adodc2 = New adodb.Recordset

    Select Case Index
     Case 0:
         Adodc2.Open "Select FAM_CODIGO,FAM_NOMBRE FROM FAMILIA", cConexCom, adOpenStatic, adLockOptimistic
         frmReferencia.Conectar Adodc2, "Select FAM_CODIGO,FAM_NOMBRE FROM FAMILIA"
         frmReferencia.Label1.Caption = "Familia de Articulos"
         frmReferencia.Show vbModal
         Adodc2.Close
         If vGUtil(1) <> "" Then
           TxF1.text = (vGUtil(1))
           'Text1(1) = VGUTIL(2)
         End If
             
   End Select
End Sub

Private Sub TxF1_GotFocus()
Enfoque TxF1
End Sub

Private Sub TxF1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then TxF1_DblClick
End Sub

Private Sub TxF1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Trim(TxF1) <> "" Then
          If existe(1, TxF1, "FAMILIA", "FAM_CODIGO", False) = False Then
            MsgBox "Codigo de Familia no existe", vbInformation, mensaje1
            TxF1.SetFocus: Exit Sub
          End If
    End If
    TxF2.SetFocus: Exit Sub
    
Else
    If KeyAscii = 39 Or KeyAscii = 34 Then KeyAscii = 0
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End If
End Sub

Private Sub TxF2_DblClick()
Static Adodc2 As adodb.Recordset
Set Adodc2 = New adodb.Recordset

    Select Case Index
     Case 0:
         Adodc2.Open "Select FAM_CODIGO,FAM_NOMBRE FROM FAMILIA", cConexCom, adOpenStatic, adLockOptimistic
         frmReferencia.Conectar Adodc2, "Select FAM_CODIGO,FAM_NOMBRE FROM FAMILIA"
         frmReferencia.Label1.Caption = "Familia de Articulos"
         frmReferencia.Show vbModal
         Adodc2.Close
         If vGUtil(1) <> "" Then
           TxF2.text = (vGUtil(1))
           'Text1(1) = VGUTIL(2)
         End If
             
   End Select
End Sub

Private Sub TxF2_GotFocus()
Enfoque TxF2
End Sub

Private Sub TxF2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then TxF2_DblClick
End Sub

Private Sub TxF2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Trim(TxF2) <> "" Then
          If existe(1, TxF2, "FAMILIA", "FAM_CODIGO", False) = False Then
            MsgBox "Codigo de Familia no existe", vbInformation, mensaje1
            TxF2.SetFocus: Exit Sub
          End If
    End If
    CmdA.SetFocus: Exit Sub
Else
    If KeyAscii = 39 Or KeyAscii = 34 Then KeyAscii = 0
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End If
End Sub

Private Sub Limpiar()
TxA1 = ""
TxA2 = ""
TxL1 = ""
TxL2 = ""
TxF1 = ""
TxF2 = ""
TxG1 = ""
TxG2 = ""
Combo1.ListIndex = 0
TxC1 = ""
TxC2 = ""
TxCon1 = ""
TxCon2 = ""
End Sub
