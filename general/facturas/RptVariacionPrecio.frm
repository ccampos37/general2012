VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form RptVariacionPrecio 
   Caption         =   "Variaciones de Precios"
   ClientHeight    =   3690
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5505
   LinkTopic       =   "Form1"
   ScaleHeight     =   3690
   ScaleWidth      =   5505
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   1560
      TabIndex        =   13
      Top             =   2250
      Width           =   915
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   1260
      TabIndex        =   7
      Top             =   2940
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   2910
      TabIndex        =   4
      Top             =   2970
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1590
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   330
      Width           =   2535
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1590
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   810
      Width           =   3615
   End
   Begin Crystal.CrystalReport oCrystalReport 
      Left            =   30
      Top             =   2130
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComCtl2.DTPicker DTHasta 
      Height          =   345
      Left            =   1590
      TabIndex        =   0
      Top             =   1770
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   40960001
      CurrentDate     =   37518
   End
   Begin MSComCtl2.DTPicker DTDesde 
      Height          =   375
      Left            =   1590
      TabIndex        =   1
      Top             =   1290
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   40960001
      CurrentDate     =   37518
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   2550
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   390
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   2910
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   840
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lbl 
      Caption         =   "Valor"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   510
      TabIndex        =   12
      Top             =   2280
      Width           =   825
   End
   Begin VB.Label lbl 
      Caption         =   "Hasta"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   510
      TabIndex        =   11
      Top             =   1770
      Width           =   615
   End
   Begin VB.Label lbl 
      Caption         =   "Desde"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   510
      TabIndex        =   10
      Top             =   1290
      Width           =   735
   End
   Begin VB.Label lbl 
      Caption         =   "Almacen"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   390
      TabIndex        =   9
      Top             =   330
      Width           =   1095
   End
   Begin VB.Label lbl 
      Caption         =   "Articulo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   390
      TabIndex        =   8
      Top             =   810
      Width           =   855
   End
End
Attribute VB_Name = "RptVariacionPrecio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adll As New dllgeneral.dll_general
Dim busca As New dll_apisgen.dll_apis
''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub cmdAceptar_Click(Index As Integer)
On Error GoTo Errores
 If DTDesde > DtHasta Then
     MsgBox "Fecha Desde debe ser mayor a Fecha Hasta", vbInformation, "AVISO"
     Exit Sub
  End If
                                   
Screen.MousePointer = 11
With oCrystalReport
        .Reset
        .ReportFileName = VGParamSistem.Rutareport & "RepVariacionPrecio.rpt"
        
'        .LogOnServer "pdssql.dll", _
'         busca.LeerIni(App.Path & "\MARFICE.INI", "BDGENERAL", "SERVIDOR", ""), _
'         busca.LeerIni(App.Path & "\MARFICE.INI", "BDGENERAL", "BDDATOS", ""), _
'         busca.LeerIni(App.Path & "\MARFICE.INI", "BDGENERAL", "USUARIO", ""), _
'         busca.LeerIni(App.Path & "\MARFICE.INI", "BDGENERAL", "PASSW", "")
'        .Connect = _
'        "DSN=" & busca.LeerIni(App.Path & "\MARFICE.INI", "BDGENERAL", "SERVIDOR", "") & ";" & _
'        "DSQ=" & busca.LeerIni(App.Path & "\MARFICE.INI", "BDGENERAL", "BDDATOS", "") & ";" & _
'        "UID=" & busca.LeerIni(App.Path & "\MARFICE.INI", "BDGENERAL", "USUARIO", "") & ";" & _
'        "PWD=" & busca.LeerIni(App.Path & "\MARFICE.INI", "BDGENERAL", "PASSW", "")
        '''''''''''''''''''''''''''''''''''''''''''''''''''
        .DiscardSavedData = True
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .WindowShowPrintSetupBtn = True
        .WindowShowExportBtn = True
        .WindowShowZoomCtl = True
        .WindowShowNavigationCtls = True
        .WindowShowPrintBtn = True
        .WindowTitle = "Variacion de Precios"
        .formulas(0) = "Empresa='" & VGParametros.nomempresa & "'"
        .formulas(1) = "Desde='" & DTDesde & "'"
        .formulas(2) = "Hasta='" & DtHasta & "'"
        If Combo1.ListIndex <> -1 Then
            .formulas(3) = "Almacen='" & Combo1.Text & "'"
        Else
            .formulas(3) = "Almacen='TODOS'"
        End If
        If Combo2.ListIndex <> -1 Then
            .formulas(4) = "Articulo='" & Combo2.Text & "'"
        Else
            .formulas(4) = "Articulo='TODOS'"
        End If
        .StoredProcParam(0) = busca.LeerIni(App.Path & "\MARFICE.INI", "CONEXION", "BDDATOS", "")
        .StoredProcParam(1) = IIf(Trim(txt(0)) = "", "%", Trim(txt(0)))
        .StoredProcParam(2) = DTDesde
        .StoredProcParam(3) = DtHasta
        .StoredProcParam(4) = IIf(Trim(txt(1)) = "", "%", Trim(txt(1)))
        .StoredProcParam(5) = IIf(Len(Trim(Text1)) = 0, 0, Text1)
        .Action = 1
  
  End With
  
Screen.MousePointer = 1

Exit Sub
Errores:
  Screen.MousePointer = 1
  MsgBox "Error inesperado: " & Err.Number & "  " & Err.Description, vbExclamation
  Err = 0
  Exit Sub
End Sub

Private Sub cmdCancelar_Click(Index As Integer)
    Unload Me
End Sub

Private Sub Combo1_Click()
  If Combo1.ListIndex <> -1 Then
    txt(0) = adll.ComboDato(Combo1.Text)
  Else
    txt(0) = ""
  End If
End Sub

Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Combo2_Click()
    If Combo2.ListIndex <> -1 Then
        txt(1) = adll.ComboDato(Combo2.Text)
    Else
        txt(1) = ""
    End If
End Sub

Private Sub Combo2_DropDown()
    
   Call adll.llenacombo _
    (Combo2, "select acodigo,RTRIM(adescri) " & _
    "from maeart inner join stkart on acodigo=stcodigo where stalma='" & Trim(txt(0)) & "' " & _
    "order by  adescri", VGCNx)
    
End Sub

Private Sub Combo2_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub DtDesde_KeyDown(KeyCode As Integer, Shift As Integer)
      If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub DtHasta_KeyDown(KeyCode As Integer, Shift As Integer)
      If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
    MostrarFormVentas Me, "C2"
    Call adll.llenacombo(Combo1, "select almacencodigo,almacendescripcion from vt_almacen", VGCNx)
    DTDesde = Date
    DtHasta = Date
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Text1 = Format(Text1, "##,##0.00")
   End If
End Sub

Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub


