VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form FrmRepMae 
   Caption         =   "Maestro de Articulo"
   ClientHeight    =   3180
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6090
   LinkTopic       =   "Form2"
   ScaleHeight     =   3180
   ScaleWidth      =   6090
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   4440
      TabIndex        =   11
      Text            =   "Text2"
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   4440
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   1320
      Width           =   735
   End
   Begin VB.OptionButton optodos 
      Caption         =   "Option1"
      Height          =   615
      Left            =   3480
      TabIndex        =   9
      Top             =   2040
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   624
      Left            =   1110
      Picture         =   "FrmRepMae.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2010
      Width           =   810
   End
   Begin VB.CommandButton Command7 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   624
      Left            =   2310
      Picture         =   "FrmRepMae.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2010
      Width           =   810
   End
   Begin VB.Frame Frame1 
      Caption         =   "Parámetros"
      Height          =   615
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Width           =   3975
      Begin VB.OptionButton OpCod 
         Caption         =   "Codigo"
         Height          =   255
         Left            =   1560
         TabIndex        =   5
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton OpDes 
         Caption         =   "Descripcion"
         Enabled         =   0   'False
         Height          =   255
         Left            =   2640
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Ordenado por "
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tipo de Consulta"
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3975
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   360
         Width           =   2100
      End
      Begin VB.Label Label1 
         Caption         =   "Por Familia"
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   360
         Width           =   1215
      End
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   300
      Top             =   2040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "FrmRepMae"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim db As Database
Dim almacen As String
Dim conexion As String
Dim Adodc3 As ADODB.Recordset

Private Sub Combo1_Click()
'almacen = Format(Combo1.ListIndex + 1, "00")
almacen = Mid(Combo1, 1, 2)
End Sub

Private Sub Combo2_Change()

End Sub

Private Sub Command7_Click()
  MousePointer = vbHourglass
  If Frame1.Visible And Frame2.Visible Then
        MousePointer = vbDefault
        Unload Me
  Else
        Frame1.Visible = True
        Frame2.Visible = True
        'FrameRep.Visible = False
        'MousePointer = vbDefault
  End If
End Sub

Private Sub Command1_Click()
   
'    If Text1 <> "" And Text2 <> "" And Text1 > Text2 Then
'        MsgBox "Ingrese un código menor al fin ", vbOKOnly, "Error"
'        Exit Sub
'    End If
'    Screen.MousePointer = 11
'    If OpArt.Value Then
'         imprimir
'    ElseIf Option2.Value Then
'        Imprimir2
'    ElseIf Option3.Value Then
'        Imprimir3
'    ElseIf Option4.Value Then
'        Imprimir4
'    ElseIf Option1.Value Then
'        Imprimir5
'    End If
'    Screen.MousePointer = 1
End Sub

Private Sub Form_Load()
Frame1.Visible = True
Frame2.Visible = True
'FrameRep.enabled = False
Carga_familia
central Me
OpCod.Value = True
'VGForm1 = 3
Combo1.ListIndex = 0
End Sub
Private Sub Carga_familia()
Dim RSQL As String
Dim rs As Recordset
Dim i As Integer
 
RSQL = "select FAM_CODIGO, FAM_NOMBRE FROM familia "
'Set db = Workspaces(0).OpenDatabase(cRuta2)
'Set RS = db.OpenRecordset(RSQL, dbOpenSnapshot)
Set rs = VGCNx.Execute(RSQL)
While Not rs.EOF
  Combo1.AddItem (rs(0)) & " - " & (rs(1))
  rs.MoveNext
Wend

Combo1.ListIndex = 1
'rs.MoveFirst
'For I = 0 To rs.RecordCount - 1
'  If rs(0) = VGAlma Then
'    Combo1.ListIndex = I
'    Exit For
'  Else
'    rs.MoveNext
'  End If
'Next
rs.Close
End Sub

Private Sub imprimir()
Dim Codigo1 As String
Dim Codigo2 As String
Dim cadena As String
Dim RSQL As String
Dim tex1 As String, tex2 As String
Dim Va1 As String, Va2 As String

''Codigo1 = UCase(Trim(Text1))
'Set Adodc3 = New ADODB.Recordset
'
'If optodos.Value Then
'    RSQL = "Select ACodigo,Adescri,ACodigo,b.STSKDIS from "
'    RSQL = RSQL & "MAEART A Inner Join STKART B on A.ACodigo=B.STCodigo "
'    RSQL = RSQL & "Where Stalma='" & almacen & "' Order by Acodigo"
'
'    CrystalReport1.Reset
'
'    Adodc3.Open RSQL, Vgcnx, adOpenDynamic, adLockOptimistic
'    If Adodc3.RecordCount > 0 Then
'      If Text1 = "" And Text2 = "" Then
'        Adodc3.MoveFirst
'        tex1 = IIf(IsNull(Adodc3("ACodigo")), "", Adodc3("ACodigo"))
'        Va1 = IIf(IsNull(Adodc3("ADescri")), "", Adodc3("Adescri"))
'        Adodc3.MoveLast
'        tex2 = IIf(IsNull(Adodc3("ACodigo")), "", Adodc3("ACodigo"))
'        Va2 = IIf(IsNull(Adodc3("ADescri")), "", Adodc3("ADescri"))
'      End If
'    End If
'    Adodc3.Close
'
'    If Check1.Value = 0 Then
'            If ChkSerie = 1 Then
'                    CrystalReport1.WindowTitle = "Inv136 -- Control de Inventarios"
'                    CrystalReport1.ReportFileName = VGParamSistem.RutaReport & IIf(Combo2.ListIndex = 1, "inv142.rpt", "inv136.rpt")
'            ElseIf Chkprecio = 0 Then
'                    CrystalReport1.WindowTitle = "Inv078 -- Control de Inventarios"
'                    CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "inv078.rpt"
'            Else
'                    CrystalReport1.WindowTitle = "Inv066-- Control de Inventarios"
'                    CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "inv066.rpt"
'            End If
'            cadena = "{STKART.STALMA}='" & almacen & "' "
'            If Chkstokcero.Value = 1 Then
'               cadena = cadena & " and {STKART.STskdis}<>0 "
'            End If
'
'    Else                                     'Consolidado
'            If ChkSerie = 1 Then
'                    CrystalReport1.WindowTitle = "Inv137 -- Control de Inventarios"
'                    CrystalReport1.ReportFileName = VGParamSistem.RutaReport & IIf(Combo2.ListIndex = 1, "inv143.rpt", "inv137.rpt")
'            ElseIf Chkprecio = 0 Then
'                    CrystalReport1.WindowTitle = "Inv079 -- Control de Inventarios"
'                    CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "inv079.rpt"
'            Else
'                    CrystalReport1.WindowTitle = "Inv081 -- Control de Inventarios"
'                    CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "inv081.rpt"
'            End If
'    End If
''    Ubi_Tab CrystalReport1
'
'    CrystalReport1.WindowShowPrintBtn = True
'    CrystalReport1.WindowShowRefreshBtn = True
'    CrystalReport1.WindowShowSearchBtn = True
'    CrystalReport1.WindowShowPrintSetupBtn = True
'    CrystalReport1.WindowState = crptMaximized
'    CrystalReport1.DiscardSavedData = True
'    CrystalReport1.Destination = crptToWindow
'


'
''    CrystalReport1.SelectionFormula = cadena
''    If Combo2.ListIndex = 1 Then
''         CrystalReport1.SortFields(0) = "+{MAEART.ADESCRI}"
''    Else
''         CrystalReport1.SortFields(0) = "+{STKART.STCODIGO}"
''    End If
'
'    CrystalReport1.Formulas(0) = "emp = '" & VGparametros.RucEmpresa & "'"
'    CrystalReport1.Formulas(1) = "almacen='" & Combo1.text & "'"
''    CrystalReport1.Formulas(1) = "hora ='" & Format(Time, "hh:mm:ss") & "'"
''    CrystalReport1.Formulas(2) = "campoini = '" & tex1 & "'"
''    CrystalReport1.Formulas(3) = "campofin = '" & tex2 & "'"
''    CrystalReport1.Formulas(4) = "detaini = '" & Va1 & "'"
''    CrystalReport1.Formulas(5) = "detafin = '" & Va2 & "'"
''
'    CrystalReport1.StoredProcParam(0) = Vgcnx.DefaultDatabase
'    CrystalReport1.StoredProcParam(1) = Left(Combo1.text, 2)
'
'    If CrystalReport1.Status <> 2 Then CrystalReport1.Action = 1
'    Exit Sub
'End If
'
'If Text1 = "" Then
'        MsgBox "Ingrese el codigo", vbExclamation, "Error"
''        Text1.SetFocus
'        Exit Sub
'End If
'  Codigo2 = UCase(Trim(Text2))
'  RSQL = "Select ADescri from MAEART Where ACodigo='" & Codigo1 & "'"
'  Adodc3.Open RSQL, Vgcnx, adOpenDynamic, adLockOptimistic
'  If Adodc3.RecordCount = 1 Then
'    Va1 = Adodc3("Adescri")
'  End If
'  Adodc3.Close
'
'  RSQL = "Select ADescri from MAEART Where ACodigo='" & Codigo2 & "'"
'  Adodc3.Open RSQL, Vgcnx, adOpenDynamic, adLockOptimistic
'  If Adodc3.RecordCount = 1 Then
'    Va2 = Adodc3("Adescri")
'  End If
'  Adodc3.Close
'
'If OpArt.Value Then           'Un select
'    If Check1.Value = 1 Then
'            If ChkSerie = 1 Then
'                    CrystalReport1.WindowTitle = "Inv137 -- Control de Inventarios"
'                    CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "inv137.rpt"
'            ElseIf Chkprecio = 0 Then
'                    CrystalReport1.WindowTitle = "Inv079 -- Control de Inventarios"
'                    CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "inv079.rpt"
'            Else
'                    CrystalReport1.WindowTitle = "Inv081 -- Control de Inventarios"
'                    CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "inv081.rpt"
'            End If
'            If Text2 <> "" Then
'                    Codigo2 = Text2
'                    cadena = "({STKART.STCODIGO} in '" & Codigo1 & "' to '" & Codigo2 & "')"
'            Else
'                    Codigo2 = Codigo1: Va2 = Va1
'                    cadena = "{STKART.STCODIGO} = '" & Codigo1 & "' "
'            End If
'    Else
'            If ChkSerie = 1 Then
'                    CrystalReport1.WindowTitle = "Inv136 -- Control de Inventarios"
'                    CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "inv136.rpt"
'            ElseIf Chkprecio = 0 Then
'                    CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "inv078.rpt"
'                    CrystalReport1.WindowTitle = "Inv078 -- Control de Inventarios"
'            Else
'                    CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "inv066.rpt"
'                    CrystalReport1.WindowTitle = "Inv066 -- Control de Inventarios"
'            End If
'            If Text2 <> "" Then
'                    Codigo2 = Text2         '  "23134671"
'                    cadena = " {STKART.STALMA}='" & almacen & "' and ({STKART.STCODIGO} in '" & Codigo1 & "' to '" & Codigo2 & "')"
'            Else
'                    Codigo2 = Codigo1: Va2 = Va1
'                    cadena = "{STKART.STALMA}='" & almacen & "' and {STKART.STCODIGO} = '" & Codigo1 & "' "
'            End If
'    End If
'    Ubi_Tab CrystalReport1
'    CrystalReport1.DiscardSavedData = True
'    CrystalReport1.Destination = crptToWindow
'    CrystalReport1.SelectionFormula = cadena
'    CrystalReport1.WindowShowPrintBtn = True
'    CrystalReport1.WindowShowRefreshBtn = True
'    CrystalReport1.WindowShowSearchBtn = True
'    CrystalReport1.WindowShowPrintSetupBtn = True
'    If Combo2.ListIndex = 1 Then
'      CrystalReport1.SortFields(0) = "+{MAEART.ADESCRI}"
'    Else
'      CrystalReport1.SortFields(0) = "+{STKART.STCODIGO}"
'    End If
'    CrystalReport1.Formulas(0) = "emp = '" & VGparametros.RucEmpresa & "'"
'    CrystalReport1.Formulas(1) = "hora ='" & Format(Time, "hh:mm:ss") & "'"
'    CrystalReport1.Formulas(2) = "campoini = '" & Codigo1 & "'"
'    CrystalReport1.Formulas(3) = "campofin = '" & Codigo2 & "'"
'    CrystalReport1.Formulas(4) = "detaini = '" & Va1 & "'"
'    CrystalReport1.Formulas(5) = "detafin = '" & Va2 & "'"
'    If CrystalReport1.Status <> 2 Then
'        CrystalReport1.Action = 1
'    End If
'End If
End Sub

Private Sub Imprimir5()
Dim Codigo1 As String
Dim Codigo2 As String
Dim cadena As String
Dim RSQL As String
Dim tex1 As String, tex2 As String
Dim Va1 As String, Va2 As String

'Codigo1 = UCase(Trim(Text1))
'Set Adodc3 = New ADODB.Recordset
'
'If optodos.Value Then
'    RSQL = "Select ACodigo,Adescri from "
'    RSQL = RSQL & "MAEART A Inner Join STKART B on A.ACodigo=B.STCodigo "
'    RSQL = RSQL & "Where Stalma='" & almacen & "' Order by Acodigo"
'
'    Adodc3.Open RSQL, Vgcnx, adOpenDynamic, adLockOptimistic
'    If Adodc3.RecordCount > 0 Then
'      If Text1 = "" And Text2 = "" Then
'        Adodc3.MoveFirst
'        tex1 = IIf(IsNull(Adodc3("ACodigo")), "", Adodc3("ACodigo"))
'        Va1 = IIf(IsNull(Adodc3("ADescri")), "", Adodc3("Adescri"))
'        Adodc3.MoveLast
'        tex2 = IIf(IsNull(Adodc3("ACodigo")), "", Adodc3("ACodigo"))
'        Va2 = IIf(IsNull(Adodc3("ADescri")), "", Adodc3("ADescri"))
'      End If
'    End If
'    Adodc3.Close
'
'    If Check1.Value = 0 Then
'            If ChkSerie = 1 Then
'                    CrystalReport1.WindowTitle = "Inv136 -- Control de Inventarios"
'                    CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "inv136.rpt"
'            ElseIf Chkprecio = 0 Then
'                    CrystalReport1.WindowTitle = "Inv138 -- Control de Inventarios"
'                    CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "inv138.rpt"
'            Else
'                    CrystalReport1.WindowTitle = "Inv140-- Control de Inventarios"
'                    CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "inv140.rpt"
'            End If
'            cadena = "{STKART.STALMA}='" & almacen & "' "
'    Else                                     'Consolidado
'            If ChkSerie = 1 Then
'                    CrystalReport1.WindowTitle = "Inv137 -- Control de Inventarios"
'                    CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "inv137.rpt"
'            ElseIf Chkprecio = 0 Then
'                    CrystalReport1.WindowTitle = "Inv139 -- Control de Inventarios"
'                    CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "inv139.rpt"
'            Else
'                    CrystalReport1.WindowTitle = "Inv140 -- Control de Inventarios"
'                    CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "inv141.rpt"
'            End If
'    End If
'    Ubi_Tab CrystalReport1
'    CrystalReport1.WindowShowPrintBtn = True
'    CrystalReport1.WindowShowRefreshBtn = True
'    CrystalReport1.WindowShowSearchBtn = True
'    CrystalReport1.WindowShowPrintSetupBtn = True
'    CrystalReport1.DiscardSavedData = True
'    CrystalReport1.Destination = crptToWindow
'    CrystalReport1.SelectionFormula = cadena
'    If Combo2.ListIndex = 1 Then
'      CrystalReport1.SortFields(0) = "+{MAEART.ADESCRI}"
'    Else
'      CrystalReport1.SortFields(0) = "+{STKART.STCODIGO}"
'    End If
'    CrystalReport1.Formulas(0) = "emp = '" & VGparametros.RucEmpresa & "'"
'    CrystalReport1.Formulas(1) = "hora ='" & Format(Time, "hh:mm:ss") & "'"
'    CrystalReport1.Formulas(2) = "campoini = '" & tex1 & "'"
'    CrystalReport1.Formulas(3) = "campofin = '" & tex2 & "'"
'    CrystalReport1.Formulas(4) = "detaini = '" & Va1 & "'"
'    CrystalReport1.Formulas(5) = "detafin = '" & Va2 & "'"
'    If CrystalReport1.Status <> 2 Then CrystalReport1.Action = 1
'    Exit Sub
'End If
'
'If Text1 = "" Then
'        MsgBox "Ingrese el codigo", vbExclamation, "Error"
'        Text1.SetFocus
'        Exit Sub
'End If
'  Codigo2 = UCase(Trim(Text2))
'  RSQL = "Select ADescri from MAEART Where ACodigo='" & Codigo1 & "'"
'  Adodc3.Open RSQL, Vgcnx, adOpenDynamic, adLockOptimistic
'  If Adodc3.RecordCount = 1 Then
'    Va1 = Adodc3("Adescri")
'  End If
'  Adodc3.Close
'
'  RSQL = "Select ADescri from MAEART Where ACodigo='" & Codigo2 & "'"
'  Adodc3.Open RSQL, Vgcnx, adOpenDynamic, adLockOptimistic
'  If Adodc3.RecordCount = 1 Then
'    Va2 = Adodc3("Adescri")
'  End If
'  Adodc3.Close
'
'If Option1.Value Then           'Un select
'    If Check1.Value = 1 Then
'            If ChkSerie = 1 Then
'                    CrystalReport1.WindowTitle = "Inv137 -- Control de Inventarios"
'                    CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "inv137.rpt"
'            ElseIf Chkprecio = 0 Then
'                    CrystalReport1.WindowTitle = "Inv139 -- Control de Inventarios"
'                    CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "inv139.rpt"
'            Else
'                    CrystalReport1.WindowTitle = "Inv141 -- Control de Inventarios"
'                    CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "inv141.rpt"
'            End If
'            If Text2 <> "" Then
'                    Codigo2 = Text2
'                    cadena = "({STKART.STCODIGO} in '" & Codigo1 & "' to '" & Codigo2 & "')"
'            Else
'                    Codigo2 = Codigo1: Va2 = Va1
'                    cadena = "{STKART.STCODIGO} = '" & Codigo1 & "' "
'            End If
'    Else
'            If ChkSerie = 1 Then
'                    CrystalReport1.WindowTitle = "Inv136 -- Control de Inventarios"
'                    CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "inv136.rpt"
'            ElseIf Chkprecio = 0 Then
'                    CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "inv138.rpt"
'                    CrystalReport1.WindowTitle = "Inv138 -- Control de Inventarios"
'            Else
'                    CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "inv040.rpt"
'                    CrystalReport1.WindowTitle = "Inv140 -- Control de Inventarios"
'            End If
'            If Text2 <> "" Then
'                    Codigo2 = Text2         '  "23134671"
'                    cadena = " {STKART.STALMA}='" & almacen & "' and ({STKART.STCODIGO} in '" & Codigo1 & "' to '" & Codigo2 & "')"
'            Else
'                    Codigo2 = Codigo1: Va2 = Va1
'                    cadena = "{STKART.STALMA}='" & almacen & "' and {STKART.STCODIGO} = '" & Codigo1 & "' "
'            End If
'    End If
'    Ubi_Tab CrystalReport1
'    CrystalReport1.DiscardSavedData = True
'    CrystalReport1.Destination = crptToWindow
'    CrystalReport1.SelectionFormula = cadena
'    CrystalReport1.WindowShowPrintBtn = True
'    CrystalReport1.WindowShowRefreshBtn = True
'    CrystalReport1.WindowShowSearchBtn = True
'    CrystalReport1.WindowShowPrintSetupBtn = True
'    If Combo2.ListIndex = 1 Then
'      CrystalReport1.SortFields(0) = "+{MAEART.ADESCRI}"
'    Else
'      CrystalReport1.SortFields(0) = "+{STKART.STCODIGO}"
'    End If
'    CrystalReport1.Formulas(0) = "emp = '" & VGparametros.RucEmpresa & "'"
'    CrystalReport1.Formulas(1) = "hora ='" & Format(Time, "hh:mm:ss") & "'"
'    CrystalReport1.Formulas(2) = "campoini = '" & Codigo1 & "'"
'    CrystalReport1.Formulas(3) = "campofin = '" & Codigo2 & "'"
'    CrystalReport1.Formulas(4) = "detaini = '" & Va1 & "'"
'    CrystalReport1.Formulas(5) = "detafin = '" & Va2 & "'"
'    If CrystalReport1.Status <> 2 Then
'        CrystalReport1.Action = 1
'    End If
'End If
End Sub

