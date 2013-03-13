VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmxxxx 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3885
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   5520
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   5520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   735
      Left            =   1692
      Picture         =   "Frmxxxx.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2952
      Width           =   735
   End
   Begin VB.CommandButton Command7 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   735
      Left            =   2892
      Picture         =   "Frmxxxx.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2952
      Width           =   735
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tipo de Consulta"
      Height          =   3768
      Left            =   36
      TabIndex        =   0
      Top             =   0
      Width           =   5415
      Begin VB.OptionButton Option2 
         Caption         =   "LOTE"
         Height          =   264
         Left            =   2988
         TabIndex        =   13
         Top             =   2592
         Width           =   1452
      End
      Begin VB.OptionButton Option1 
         Caption         =   "SERIE"
         Height          =   264
         Left            =   1440
         TabIndex        =   12
         Top             =   2592
         Value           =   -1  'True
         Width           =   1452
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Consolidado"
         Height          =   225
         Left            =   2016
         TabIndex        =   11
         Top             =   2124
         Width           =   1335
      End
      Begin VB.ComboBox Combo2 
         Height          =   288
         ItemData        =   "Frmxxxx.frx":0884
         Left            =   2052
         List            =   "Frmxxxx.frx":088E
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1692
         Width           =   1440
      End
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2040
         MaxLength       =   20
         TabIndex        =   6
         Top             =   1260
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   2040
         MaxLength       =   20
         TabIndex        =   5
         Top             =   792
         Width           =   1470
      End
      Begin VB.ComboBox Combo1 
         Height          =   288
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   360
         Width           =   3180
      End
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         PrintFileLinesPerPage=   60
      End
      Begin VB.Label Label6 
         Caption         =   "Ordenado por "
         Height          =   252
         Left            =   972
         TabIndex        =   10
         Top             =   1704
         Width           =   996
      End
      Begin VB.Label Label3 
         Caption         =   "Fin"
         Height          =   252
         Left            =   1440
         TabIndex        =   8
         Top             =   1296
         Width           =   732
      End
      Begin VB.Label Label2 
         Caption         =   "Inicio "
         Height          =   252
         Left            =   1368
         TabIndex        =   7
         Top             =   804
         Width           =   732
      End
      Begin VB.Label Label1 
         Caption         =   "Por Almacen"
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   360
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmxxxx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim db As Database
Dim almacen As String
Dim Conexion As String
Dim Adodc3 As ADODB.Recordset

Private Sub Combo1_Click()
'almacen = Format(Combo1.ListIndex + 1, "00")
almacen = Mid(Combo1, 1, 2)
End Sub

Private Sub Command1_Click()
If Option1.Value = True Then
    CrystalReport1.WindowTitle = "Inv511 -- Control de Inventarios"
    CrystalReport1.ReportFileName = cRutP & IIf(Combo2.ListIndex = 1, "inv142.rpt", "inv511.rpt")
    
    Ubi_Tab CrystalReport1
    CrystalReport1.WindowShowPrintBtn = True
    CrystalReport1.WindowShowRefreshBtn = True
    CrystalReport1.WindowShowSearchBtn = True
    CrystalReport1.WindowShowPrintSetupBtn = True
    CrystalReport1.DiscardSavedData = True
    CrystalReport1.Destination = crptToWindow
    'CrystalReport1.SelectionFormula = cadena
    If Combo2.ListIndex = 1 Then
         CrystalReport1.SortFields(0) = "+{MAEART.ADESCRI}"
    Else
         CrystalReport1.SortFields(0) = "+{MAEART.ACODIGO}"
    End If
    
    If Text1 = "" Or Text2 = "" Then
       If Check1.Value = 1 Then
          CrystalReport1.ReplaceSelectionFormula ("{STKSERI.STSSKDIS}<>0 ")
          CrystalReport1.Formulas(0) = "xocultaGrupoalma=false "
       Else
          CrystalReport1.ReplaceSelectionFormula ("{STKSERI.STSSKDIS}<>0  AND {STKSERI.STSALMA}='" & VGAlma & "'")
          CrystalReport1.Formulas(0) = "xocultaGrupoalma=true"
       End If
    Else
       If Check1.Value = 1 Then
          CrystalReport1.ReplaceSelectionFormula ("{STKSERI.STSSKDIS}<>0 and ( {MAEART.ACODIGO}>='" & Text1 & "' and {MAEART.ACODIGO}<='" & Text2 & "' ) AND {MAEART.ACODIGO}<>'' ")
          CrystalReport1.Formulas(0) = "xocultaGrupoalma=false "
       Else
          CrystalReport1.ReplaceSelectionFormula ("{STKSERI.STSSKDIS}<>0 and ( {MAEART.ACODIGO}>='" & Text1 & "' and {MAEART.ACODIGO}<='" & Text2 & "' ) AND {MAEART.ACODIGO}<>'' AND {STKSERI.STSALMA}='" & VGAlma & "'")
          CrystalReport1.Formulas(0) = "xocultaGrupoalma=true"
       End If
    End If
 
    CrystalReport1.Formulas(1) = "emp = '" & VGNemp & "'"
    CrystalReport1.Formulas(2) = "hora ='" & Format(Time, "hh:mm:ss") & "'"
    CrystalReport1.Formulas(3) = "campoini = '" & Text1 & "'"
    CrystalReport1.Formulas(4) = "campofin = '" & Text2 & "'"
    'CrystalReport1.Formulas(4) = "detaini = '" & Va1 & "'"
    'CrystalReport1.Formulas(5) = "detafin = '" & Va2 & "'"
'    If CrystalReport1.Status <> 2 Then CrystalReport1.Action = 1
    Exit Sub
 Else
    CrystalReport1.WindowTitle = "Inv512 -- Control de Inventarios"
    CrystalReport1.ReportFileName = cRutP & "inv512.rpt"
    
    Ubi_Tab CrystalReport1
    CrystalReport1.WindowShowPrintBtn = True
    CrystalReport1.WindowShowRefreshBtn = True
    CrystalReport1.WindowShowSearchBtn = True
    CrystalReport1.WindowShowPrintSetupBtn = True
    CrystalReport1.DiscardSavedData = True
    CrystalReport1.Destination = crptToWindow
    'CrystalReport1.SelectionFormula = cadena
    If Combo2.ListIndex = 1 Then
         CrystalReport1.SortFields(0) = "+{MAEART.ADESCRI}"
    Else
         CrystalReport1.SortFields(0) = "+{MAEART.ACODIGO}"
    End If
    If Check1.Value = 1 Then
       CrystalReport1.ReplaceSelectionFormula ("{STKLOTE.STSLKDIS}<>0 and ( {MAEART.ACODIGO}>='" & Text1 & "' and {MAEART.ACODIGO}<='" & Text2 & "' ) AND {MAEART.ACODIGO}<>'' ")
       CrystalReport1.Formulas(0) = "xocultaGrupoalma=false "
    Else
       CrystalReport1.ReplaceSelectionFormula ("{STKLOTE.STSLKDIS}<>0 and ( {MAEART.ACODIGO}>='" & Text1 & "' and {MAEART.ACODIGO}<='" & Text2 & "' ) AND {MAEART.ACODIGO}<>'' AND {STKLOTE.STSALMA}='" & VGAlma & "'")
       CrystalReport1.Formulas(0) = "xocultaGrupoalma=true"
    End If
 
    CrystalReport1.Formulas(1) = "emp = '" & VGNemp & "'"
    CrystalReport1.Formulas(2) = "hora ='" & Format(Time, "hh:mm:ss") & "'"
    CrystalReport1.Formulas(3) = "campoini = '" & Text1 & "'"
    CrystalReport1.Formulas(4) = "campofin = '" & Text2 & "'"
    'CrystalReport1.Formulas(4) = "detaini = '" & Va1 & "'"
    'CrystalReport1.Formulas(5) = "detafin = '" & Va2 & "'"
    If CrystalReport1.Status <> 2 Then CrystalReport1.Action = 1
    Exit Sub
 End If
End Sub

Private Sub Command7_Click()
Unload Me
End Sub

Private Sub Form_Load()
Carga_Almacen
central Me
Combo2.ListIndex = 0
End Sub
Private Sub Carga_Almacen()
Dim RSQL As String
Dim rs As Recordset
Dim i As Integer
 
RSQL = "select TAALMA,TADESCRI FROM TabAlm "
'Set db = Workspaces(0).OpenDatabase(cRuta2)
'Set RS = db.OpenRecordset(RSQL, dbOpenSnapshot)
Set rs = cConexCom.Execute(RSQL)
While Not rs.EOF
  Combo1.AddItem (rs(0)) & "  " & (rs(1))
  rs.MoveNext
Wend
'Combo1.ListIndex = CInt(VGAlma) - 1
rs.MoveFirst
For i = 0 To rs.RecordCount - 1
  If rs(0) = VGAlma Then
    Combo1.ListIndex = i
    Exit For
  Else
    rs.MoveNext
  End If
Next
rs.Close
End Sub

Private Sub Text1_DblClick()
Dim Adodc2 As ADODB.Recordset
Set Adodc2 = New ADODB.Recordset
   
   VGForm1 = 17
   FormAyuArt1.Show 1
   If Text1 <> "" And Text2 <> "" And Text1 > Text2 Then
        MsgBox "Ingrese un codigo menor al fin ", vbOKOnly, "Error"
        Exit Sub
   End If
   If Text1 <> "" Then
        Text2.Enabled = True
        Text2.SetFocus
   End If
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then Text1_DblClick
End Sub

Private Sub Text2_DblClick()
Dim Adodc2 As ADODB.Recordset
Set Adodc2 = New ADODB.Recordset
   
   VGForm1 = 17
   FormAyuArt1.Show 1
   If Text2 <> "" Then
        Command1.SetFocus
   End If
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then Text2_DblClick
End Sub
