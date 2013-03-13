VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmStockLoteSerie 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Stock de Articulos por Serie/Lote"
   ClientHeight    =   3795
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   5505
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   5505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   735
      Left            =   1935
      Picture         =   "frmStockLoteSerie.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2715
      Width           =   735
   End
   Begin VB.CommandButton Command7 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   735
      Left            =   3135
      Picture         =   "frmStockLoteSerie.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2715
      Width           =   735
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tipo de Consulta"
      Height          =   3765
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   5415
      Begin VB.OptionButton Option2 
         Caption         =   "LOTE"
         Height          =   264
         Left            =   3105
         TabIndex        =   12
         Top             =   2235
         Value           =   -1  'True
         Width           =   1452
      End
      Begin VB.OptionButton Option1 
         Caption         =   "SERIE"
         Enabled         =   0   'False
         Height          =   264
         Left            =   1560
         TabIndex        =   11
         Top             =   2235
         Width           =   1452
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "frmStockLoteSerie.frx":0884
         Left            =   2055
         List            =   "frmStockLoteSerie.frx":088E
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1695
         Visible         =   0   'False
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
         Height          =   255
         Left            =   975
         TabIndex        =   10
         Top             =   1710
         Visible         =   0   'False
         Width           =   990
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
         Left            =   840
         TabIndex        =   2
         Top             =   360
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmStockLoteSerie"
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
Dim c_alma As String
Dim c_flagpro As Integer
Dim descri As String
Dim puntero As Integer
  If Trim(Combo1.text) = "" Then
      Combo1.SetFocus
      Exit Sub
    
  Else
      c_alma = Left(Combo1.text, 2)
      'puntero = InStr(Combo1.text, "-")
      descri = Right(Combo1.text, Len(Combo1.text) - puntero)
  End If

If Text1 <> "" And Text2 <> "" Then
    c_flagpro = 1
Else
    c_flagpro = 0
End If

If Option1.Value = True Then
    CrystalReport1.WindowTitle = "Inv511 -- Control de Inventarios"
    CrystalReport1.ReportFileName = VGParamSistem.RutaReport & IIf(Combo2.ListIndex = 1, "inv142.rpt", "inv511.rpt")
    
'     CrystalReport1.Connect = VGcadenareport2
      
    CrystalReport1.WindowShowPrintBtn = True
    CrystalReport1.WindowShowRefreshBtn = True
    CrystalReport1.WindowShowSearchBtn = True
    CrystalReport1.WindowShowPrintSetupBtn = True
    CrystalReport1.DiscardSavedData = True
    CrystalReport1.Destination = crptToWindow
    
    
    
    If Combo2.ListIndex = 1 Then
         CrystalReport1.SortFields(0) = "+{MAEART.ADESCRI}"
    Else
         CrystalReport1.SortFields(0) = "+{MAEART.ACODIGO}"
    End If
    

 
    CrystalReport1.formulas(1) = "emp = '" & VGparametros.RucEmpresa & "'"
    CrystalReport1.formulas(2) = "hora ='" & Format(Time, "hh:mm:ss") & "'"
    CrystalReport1.formulas(3) = "campoini = '" & Text1 & "'"
    CrystalReport1.formulas(4) = "campofin = '" & Text2 & "'"
    'CrystalReport1.Formulas(4) = "detaini = '" & Va1 & "'"
    'CrystalReport1.Formulas(5) = "detafin = '" & Va2 & "'"
'    If CrystalReport1.Status <> 2 Then CrystalReport1.Action = 1
    Exit Sub
 Else
    
    CrystalReport1.Reset
    CrystalReport1.WindowTitle = "Inv512 -- Control de Inventarios"
    CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "inv512.rpt"
    
       
        If VGsql = 1 Then
           CrystalReport1.Connect = "Provider=SQLOLEDB;Password=" & VGParamSistem.PwdGEN & ";Persist Security Info=True;User ID=" & VGParamSistem.UsuarioGEN & ";Initial Catalog=" & VGParamSistem.BDEmpresaGEN & ";SERVER=" & VGParamSistem.ServidorGEN
          Else
           CrystalReport1.Connect = VGcadenareport2
           CrystalReport1.LogOnServer "pdssql.dll", "", VGParamSistem.BDEmpresaGEN, VGParamSistem.UsuarioGEN, VGParamSistem.PwdGEN

        End If
      CrystalReport1.WindowShowPrintBtn = True
      CrystalReport1.WindowShowRefreshBtn = True
      CrystalReport1.WindowShowSearchBtn = True
      CrystalReport1.WindowShowPrintSetupBtn = True
      CrystalReport1.DiscardSavedData = True
      CrystalReport1.Destination = crptToWindow
      CrystalReport1.WindowState = crptMaximized
      
       
       CrystalReport1.StoredProcParam(0) = CStr(VGCNx.DefaultDatabase)
       CrystalReport1.StoredProcParam(1) = c_alma
       CrystalReport1.StoredProcParam(2) = c_flagpro
       CrystalReport1.StoredProcParam(3) = IIf(Trim(Text1.text) = "", "%%", Trim(Text1.text))
       CrystalReport1.StoredProcParam(4) = IIf(Trim(Text2.text) = "", "%%", Trim(Text2.text))
       
   
       If Combo2.ListIndex = 1 Then
            'CrystalReport1.SortFields(0) = {stcodigo}
       Else
            'CrystalReport1.SortFields(0) = "+{MAEART.ACODIGO}"
       End If

       CrystalReport1.formulas(0) = "alma = '" & descri & "'"
       CrystalReport1.formulas(1) = "emp = '" & VGparametros.RucEmpresa & "' "
       CrystalReport1.formulas(2) = "ini = '" & Text1 & "'"
       CrystalReport1.formulas(3) = "fin = '" & Text2 & "'"
       
       'Resume
       If CrystalReport1.Status <> 2 Then CrystalReport1.Action = 1
       Screen.MousePointer = 1
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
Dim rsql As String
Dim rs As Recordset
Dim I As Integer
 
rsql = "select TAALMA,TADESCRI FROM TabAlm "
'Set db = Workspaces(0).OpenDatabase(cRuta2)
'Set RS = db.OpenRecordset(RSQL, dbOpenSnapshot)
Set rs = VGCNx.Execute(rsql)
While Not rs.EOF
  Combo1.AddItem (rs(0)) & "  " & (rs(1))
  rs.MoveNext
Wend
'Combo1.ListIndex = CInt(VGAlma) - 1
rs.MoveFirst
For I = 0 To rs.RecordCount - 1
  If rs(0) = VGAlma Then
    Combo1.ListIndex = I
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
