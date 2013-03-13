VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form FrmLisArticulos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Articulos"
   ClientHeight    =   5610
   ClientLeft      =   3855
   ClientTop       =   1485
   ClientWidth     =   3720
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   3720
   Begin VB.Frame Frame9 
      Height          =   1095
      Left            =   48
      TabIndex        =   39
      Top             =   4485
      Width           =   3648
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
      Height          =   2910
      Left            =   96
      TabIndex        =   0
      Top             =   15
      Width           =   3576
      Begin VB.OptionButton Option1 
         Caption         =   "Por Tipo Ubicación"
         Enabled         =   0   'False
         Height          =   375
         Index           =   8
         Left            =   495
         TabIndex        =   44
         Top             =   2490
         Width           =   2055
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Por Tipo Articulo"
         Enabled         =   0   'False
         Height          =   375
         Index           =   7
         Left            =   480
         TabIndex        =   22
         Top             =   2175
         Width           =   2055
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Por Cta. Contable"
         Enabled         =   0   'False
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
         Top             =   1110
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
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
   End
   Begin VB.Frame Frame6 
      Height          =   1425
      Left            =   48
      TabIndex        =   19
      Top             =   3120
      Visible         =   0   'False
      Width           =   3612
      Begin VB.ComboBox Combo1 
         Height          =   288
         ItemData        =   "FrmLisArticulos.frx":0884
         Left            =   816
         List            =   "FrmLisArticulos.frx":0891
         Style           =   2  'Dropdown List
         TabIndex        =   42
         Top             =   576
         Width           =   2616
      End
      Begin VB.CheckBox Check7 
         Caption         =   "Listar Todo"
         Height          =   255
         Left            =   192
         TabIndex        =   38
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "Tipo   :"
         Height          =   252
         Left            =   216
         TabIndex        =   43
         Top             =   648
         Width           =   744
      End
   End
   Begin VB.Frame Frame5 
      Height          =   1395
      Left            =   144
      TabIndex        =   14
      Top             =   3216
      Visible         =   0   'False
      Width           =   3576
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
      Height          =   1425
      Left            =   96
      TabIndex        =   23
      Top             =   3120
      Visible         =   0   'False
      Width           =   3576
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
         Height          =   252
         Left            =   96
         TabIndex        =   26
         Top             =   624
         Width           =   1452
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1470
      Left            =   48
      TabIndex        =   4
      Top             =   3216
      Visible         =   0   'False
      Width           =   3576
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
      Height          =   1410
      Left            =   48
      TabIndex        =   9
      Top             =   3216
      Visible         =   0   'False
      Width           =   3576
      Begin VB.CheckBox Check3 
         Caption         =   "Listar Todo"
         Height          =   255
         Left            =   240
         TabIndex        =   33
         Top             =   288
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
      Height          =   1530
      Left            =   96
      TabIndex        =   28
      Top             =   2928
      Visible         =   0   'False
      Width           =   3576
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
Option Explicit
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
    If Existe(1, TxA1, "MAeart", "Acodigo", False) = False Then
        MsgBox "Código no Existe", vbInformation, "Mensaje"
        TxA1.SetFocus
        Exit Sub
    End If
    If Existe(1, TxA2, "MAeart", "Acodigo", False) = False Then
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
    If Check1.Value = 1 Then
        nOpc = 3
        ImpT
        Exit Sub
    End If
    If TxF1 <> "" Then
    If fFam(TxF1) = "" Then
        MsgBox "Código de familia no Existe", vbInformation, "Mensaje"
        TxF1.SetFocus
        Exit Sub
    End If
    End If
    If TxF2 <> "" Then
    If fFam(TxF2) = "" Then
        MsgBox "Código de familia  no Existe", vbInformation, "Mensaje"
        TxF2.SetFocus
        Exit Sub
    End If
    End If
    
    If Trim(TxF1) > Trim(TxF2) Then
        MsgBox "La Familia Inicial es Mayor a la Familia Final"
        TxF1.SetFocus
        Exit Sub
    Else
        nOpc = 3
        ImpT
    End If
    
ElseIf Option1(8).Value = True Then 'Descripcion
    nOpc = 8
    ImpT
ElseIf Option1(3).Value = True Then 'Familia
    If Check1.Value = 1 Then
        nOpc = 3
        ImpT
        Exit Sub
    End If
    If TxF1 <> "" Then
    If fFam(TxF1) = "" Then
        MsgBox "Código de familia no Existe", vbInformation, "Mensaje"
        TxF1.SetFocus
        Exit Sub
    End If
    End If
    If TxF2 <> "" Then
    If fFam(TxF2) = "" Then
        MsgBox "Código de Familia no Existe", vbInformation, "Mensaje"
        TxF2.SetFocus
        Exit Sub
    End If
    End If
    
    If Trim(TxF1) > Trim(TxF2) Then
        MsgBox "La Familia Inicial es Mayor a la Familia Final"
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
    If Not (Existe(1, TxC1, "MAeArt", "ACODIGO2 ", False)) Then
        MsgBox "Código no Existe", vbInformation, "Mensaje"
        TxC1.SetFocus
        Exit Sub
    End If
    End If
    If TxC2 <> "" Then
    If Not (Existe(1, TxC2, "MAeArt", "ACODIGO2 ", False)) Then
        MsgBox "Código no Existe", vbInformation, "Mensaje"
        TxC2.SetFocus
        Exit Sub
    End If
    End If
    
    If Trim(TxC1) > Trim(TxC2) Then
        MsgBox "El Cod. Fabricante Inicial es Mayor al Cod. Fabr. Final Final"
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
    
    If Devolver_Dato(3, TxCon1, "PLAN_CUENTA_NACIONAL ", "PLANCTA_CODIGO", False, "PLANCTA_DESCRIPCION") = "" Then
        MsgBox "Código no Existe", vbInformation, "Mensaje"
        TxCon1.SetFocus
        Exit Sub
    End If
    End If
    
    If TxCon2 <> "" Then
    If Devolver_Dato(3, TxCon2, "PLAN_CUENTA_NACIONAL ", "PLANCTA_CODIGO", False, "PLANCTA_DESCRIPCION") = "" Then
        MsgBox "Código no Existe", vbInformation, "Mensaje"
        TxCon2.SetFocus
        Exit Sub
    End If
    End If
    
    If Trim(TxCon1) > Trim(TxCon2) Then
        MsgBox "La Cta. Cont. Inicial es Mayor al Cta.Cont. Final"
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

CrystalReport1.Reset
CrystalReport1.Destination = crptToWindow
CrystalReport1.WindowShowPrintBtn = True
CrystalReport1.WindowShowRefreshBtn = True
CrystalReport1.WindowShowSearchBtn = True
CrystalReport1.WindowShowPrintSetupBtn = True
CrystalReport1.WindowState = crptMaximized
CrystalReport1.WindowTitle = "al_catalogoarticulo -- Sistema de Inventarios "
CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "al_catalogoarticulo.Rpt" '
If VGsql = 1 Then
    CrystalReport1.Connect = "Provider=SQLOLEDB;Password=" & VGParamSistem.PwdGEN & ";Persist Security Info=True;User ID=" & VGParamSistem.UsuarioGEN & ";Initial Catalog=" & VGParamSistem.BDEmpresaGEN & ";SERVER=" & VGParamSistem.ServidorGEN
 Else
    CrystalReport1.Connect = VGcadenareport2
End If
CrystalReport1.StoredProcParam(0) = VGCNx.DefaultDatabase
CrystalReport1.formulas(0) = "emp ='" & VGparametros.NomEmpresa & "'"
If nOpc = 0 Then  'Codigo
    If Check4.Value = 1 Then
        CrystalReport1.ReplaceSelectionFormula ("")
    Else
        CrystalReport1.ReplaceSelectionFormula ("{al_listaarticulo_rep.ACODIGO} >=  '" & Trim(TxA1) & "' and {al_listaarticulo_rep.ACODIGO} <=  '" & Trim(TxA2) & "'")
    End If
    If CrystalReport1.Status <> 2 Then CrystalReport1.Action = 1
    
ElseIf nOpc = 1 Then  'Descripcion
    If Check1.Value = 1 Then
        CrystalReport1.ReplaceSelectionFormula ("")
    Else
        If TxF1.text = "" Or TxF2.text = "" Then Exit Sub
        CrystalReport1.ReplaceSelectionFormula ("{al_listaarticulo_rep.AFAMILIA} >=  '" & Trim(TxF1) & "' and {al_listaarticulo_rep.AFAMILIA} <=  '" & Trim(TxF2) & "'")
    End If
    If CrystalReport1.Status <> 2 Then CrystalReport1.Action = 1
ElseIf nOpc = 3 Then
    If Check1.Value = 1 Then
        CrystalReport1.ReplaceSelectionFormula ("")
    Else
        If TxF1.text = "" Or TxF2.text = "" Then Exit Sub
        CrystalReport1.ReplaceSelectionFormula ("{al_listaarticulo_rep.AFAMILIA} >=  '" & Trim(TxF1) & "' and {al_listaarticulo_rep.AFAMILIA} <=  '" & Trim(TxF2) & "'")
    End If
    If CrystalReport1.Status <> 2 Then CrystalReport1.Action = 1
ElseIf nOpc = 5 Then
    Call Ubi_Tab(CrystalReport1)
    If Check6.Value = 1 Then
        CrystalReport1.ReplaceSelectionFormula ("")
    Else
        CrystalReport1.ReplaceSelectionFormula ("{al_listaarticulo_rep.ACODIGO2} >=  '" & Trim(TxC1) & "' and {al_listaarticulo_rep.ACODIGO2} <=  '" & Trim(TxC2) & "'")
    End If
    If CrystalReport1.Status <> 2 Then CrystalReport1.Action = 1

ElseIf nOpc = 6 Then
    '
    If Check5.Value = 1 Then
        CrystalReport1.ReplaceSelectionFormula ("")
    Else
        If TxCon1.text = "" Or TxCon2.text = "" Then Exit Sub
        CrystalReport1.ReplaceSelectionFormula ("{al_listaarticulo_rep.ACUENTA} >=  '" & Trim(TxCon1) & "' and {al_listaarticulo_rep.ACUENTA} <=  '" & Trim(TxCon2) & "'")
    End If
    If CrystalReport1.Status <> 2 Then CrystalReport1.Action = 1
ElseIf nOpc = 7 Then
    CrystalReport1.WindowTitle = "Inv014 -- Sistema de Inventarios"
    CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "inv014.Rpt"
    Call Ubi_Tab(CrystalReport1)
    If Check7.Value = 1 Then
        CrystalReport1.ReplaceSelectionFormula ("")
    Else
        If Combo1.ListIndex = 0 Then
        CrystalReport1.ReplaceSelectionFormula ("{al_listaarticulo_rep.ATIPO} =  '" & Left(Combo1.text, 2) & "'")
        ElseIf Combo1.ListIndex = 1 Then
               CrystalReport1.ReplaceSelectionFormula ("{al_listaarticulo_rep.ATIPO} =  'N'")
             ElseIf Combo1.ListIndex = 2 Then
                   CrystalReport1.ReplaceSelectionFormula ("{al_listaarticulo_rep.ATIPO} =  'S'")
        End If
    End If
    If CrystalReport1.Status <> 2 Then CrystalReport1.Action = 1
ElseIf nOpc = 8 Then
    If Check4.Value = 1 Then
        CrystalReport1.ReplaceSelectionFormula ("")
    Else
        CrystalReport1.ReplaceSelectionFormula ("{al_listaarticulo_rep.ACODIGO} >=  '" & Trim(TxA1) & "' and {al_listaarticulo_rep.ACODIGO} <=  '" & Trim(TxA2) & "'")
    End If
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

Call CargaTipo

Frame2.Visible = True
Frame3.Visible = False
'Frame4.Visible = False
Frame5.Visible = False
Frame6.Visible = False
Frame7.Visible = False
Frame8.Visible = False
Option1(0).Value = True

CTIME = Format(Time, "hh:mm:ss")
CrystalReport1.WindowTitle = "Sistema  de Inventarios"
CrystalReport1.formulas(0) = "Hora = '" & CTIME & "'"
CrystalReport1.formulas(1) = "Empresa = '" & Mid(VGparametros.RucEmpresa, 1, 20) & "'"
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
    Frame5.Visible = True
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
Case 8 ' Articulo
    Frame2.Visible = True
    Frame3.Visible = False
    'Frame4.Visible = False
    Frame5.Visible = False
    Frame6.Visible = False
    Frame7.Visible = False
    Frame8.Visible = False
End Select
Limpiar
End Sub

Private Sub TxA1_DblClick()
Static Adodc2 As ADODB.Recordset
Set Adodc2 = New ADODB.Recordset
    
Adodc2.Open "SELECT ACODIGO,ADESCRI FROM maeart", VGCNx, adOpenStatic, adLockOptimistic
frmReferencia.Conectar Adodc2, "SELECT ACODIGO,ADESCRI FROM MAEART"
frmReferencia.Label1.Caption = "Articulos"
frmReferencia.Show vbModal
Adodc2.Close
If vGUtil(1) <> "" Then
  TxA1.text = (vGUtil(1))
End If
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
          If Existe(1, TxA1, "MAEART", "ACODIGO", False) = False Then
            MsgBox "Codigo de Articulo no existe", vbInformation, "Inventarios"
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

Adodc2.Open "SELECT ACODIGO,ADESCRI FROM MAEART", VGCNx, adOpenStatic, adLockOptimistic
frmReferencia.Conectar Adodc2, "SELECT ACODIGO,ADESCRI FROM MAEART"
frmReferencia.Label1.Caption = "Articulos"
frmReferencia.Show vbModal
Adodc2.Close
If vGUtil(1) <> "" Then
  TxA2.text = (vGUtil(1))
End If
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
          If Existe(1, TxA2, "MAEART", "ACODIGO", False) = False Then
            MsgBox "Codigo de Articulo no existe", vbInformation, "Inventarios"
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

Adodc2.Open "Select  ACODIGO2,ADESCRI from MAEART where ACODIGO2<>''", VGCNx, adOpenStatic, adLockOptimistic
frmReferencia.Conectar Adodc2, "Select  ACODIGO2,ADESCRI from MAEART where ACODIGO2<>''"
frmReferencia.Label1.Caption = "Codigo de Fabricantes"
frmReferencia.Show vbModal
Adodc2.Close
If vGUtil(1) <> "" Then
  TxC1.text = (vGUtil(1))
End If
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
          If Existe(1, TxC1, "MAEART", "ACODIGO2", False) = False Then
             MsgBox "Codigo de Proveedor no existe", vbInformation, "Inventarios"
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
Static Adodc2 As ADODB.Recordset
Set Adodc2 = New ADODB.Recordset

Adodc2.Open "Select  ACODIGO2,ADESCRI  from MAEART where ACODIGO2<>''", VGCNx, adOpenStatic, adLockOptimistic
frmReferencia.Conectar Adodc2, "Select  ACODIGO2,ADESCRI  from MAEART where ACODIGO2<>''"
frmReferencia.Label1.Caption = "Codigo de Fabricantes"
frmReferencia.Show vbModal
Adodc2.Close
If vGUtil(1) <> "" Then
  TxC2.text = (vGUtil(1))
End If
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
          If Existe(1, TxC2, "MAEART", "ACODIGO2", False) = False Then
            MsgBox "Codigo de Proveedor no existe", vbInformation, "Inventarios"
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
Static Adodc2 As ADODB.Recordset
Set Adodc2 = New ADODB.Recordset
On Local Error GoTo ERRAR

Adodc2.Open "Select PLANCTA_CODIGO,PLANCTA_DESCRIPCION FROM PLAN_CUENTA_NACIONAL", VGcnxCT, adOpenStatic, adLockOptimistic
frmReferencia.Conectar Adodc2, "Select PLANCTA_CODIGO,PLANCTA_DESCRIPCION FROM PLAN_CUENTA_NACIONAL"
frmReferencia.Label1.Caption = "Cuenta Contable"
frmReferencia.Show vbModal
Adodc2.Close
If vGUtil(1) <> "" Then
  TxCon1.text = (vGUtil(1))
End If
Exit Sub
ERRAR:

MsgBox "No Tiene Enlace con Contabilidad.........!", vbInformation, "Aviso"

Exit Sub
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
          If Existe(1, TxCon1, "PLAN_CUENTA_NACIONAL", "PLANCTA_CODIGO", False) = False Then
            MsgBox "Codigo de Cuenta no existe", vbInformation, "Inventarios"
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
Static Adodc2 As ADODB.Recordset
Set Adodc2 = New ADODB.Recordset
On Local Error GoTo ERRAR
Adodc2.Open "Select PLANCTA_CODIGO,PLANCTA_DESCRIPCION FROM PLAN_CUENTA_NACIONAL", VGcnxCT, adOpenStatic, adLockOptimistic
frmReferencia.Conectar Adodc2, "Select PLANCTA_CODIGO,PLANCTA_DESCRIPCION FROM PLAN_CUENTA_NACIONAL"
frmReferencia.Label1.Caption = "Cuenta Contable"
frmReferencia.Show vbModal
Adodc2.Close
If vGUtil(1) <> "" Then
  TxCon2.text = (vGUtil(1))
End If

Exit Sub
ERRAR:

MsgBox "No Tiene Enlace con Contabilidad.........!", vbInformation, "Aviso"

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
          If Existe(1, TxCon2, "CUENTA_CONTABLE", "COD_CUENTA", False) = False Then
            MsgBox "Codigo de Cuenta no existe", vbInformation, "Inventarios"
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
Static Adodc2 As ADODB.Recordset
Set Adodc2 = New ADODB.Recordset

Adodc2.Open "Select FAM_CODIGO,FAM_NOMBRE FROM FAMILIA", VGCNx, adOpenStatic, adLockOptimistic
frmReferencia.Conectar Adodc2, "Select FAM_CODIGO,FAM_NOMBRE FROM FAMILIA"
frmReferencia.Label1.Caption = "Familia de Articulos"
frmReferencia.Show vbModal
Adodc2.Close
If vGUtil(1) <> "" Then
  TxF1.text = (vGUtil(1))
End If
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
          If Existe(1, TxF1, "FAMILIA", "FAM_CODIGO", False) = False Then
            MsgBox "Codigo de Familia no existe", vbInformation, "Inventarios"
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
Static Adodc2 As ADODB.Recordset
Set Adodc2 = New ADODB.Recordset

Adodc2.Open "Select FAM_CODIGO,FAM_NOMBRE FROM FAMILIA", VGCNx, adOpenStatic, adLockOptimistic
frmReferencia.Conectar Adodc2, "Select FAM_CODIGO,FAM_NOMBRE FROM FAMILIA"
frmReferencia.Label1.Caption = "Familia de Articulos"
frmReferencia.Show vbModal
Adodc2.Close
If vGUtil(1) <> "" Then
  TxF2.text = (vGUtil(1))
End If
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
          If Existe(1, TxF2, "FAMILIA", "FAM_CODIGO", False) = False Then
            MsgBox "Codigo de Familia no existe", vbInformation, "Inventarios"
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
TxF1 = ""
TxF2 = ""
TxG1 = ""
TxG2 = ""
'Combo1.ListIndex = 0
TxC1 = ""
TxC2 = ""
TxCon1 = ""
TxCon2 = ""
End Sub

Sub CargaTipo()
Dim rs As New ADODB.Recordset

End Sub
