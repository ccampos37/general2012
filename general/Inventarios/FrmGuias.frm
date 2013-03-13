VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmGuias 
   Caption         =   "Guías de Remisión"
   ClientHeight    =   4050
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6330
   Icon            =   "FrmGuias.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   6330
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   750
      Left            =   240
      TabIndex        =   15
      Top             =   60
      Width           =   5820
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "FrmGuias.frx":08CA
         Left            =   2325
         List            =   "FrmGuias.frx":08D4
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   2100
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo de Guías        :"
         Height          =   255
         Left            =   645
         TabIndex        =   16
         Top             =   255
         Width           =   1515
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Rango de Fecha"
      Height          =   1965
      Left            =   3405
      TabIndex        =   12
      Top             =   1035
      Width           =   2670
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   300
         Left            =   1050
         TabIndex        =   6
         Top             =   1095
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   529
         _Version        =   393216
         Format          =   24838145
         CurrentDate     =   36754
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   300
         Left            =   1050
         TabIndex        =   5
         Top             =   525
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   529
         _Version        =   393216
         Format          =   24838145
         CurrentDate     =   36754
      End
      Begin VB.Label Label4 
         Caption         =   "Inicio"
         Height          =   225
         Left            =   300
         TabIndex        =   14
         Top             =   585
         Width           =   660
      End
      Begin VB.Label Label5 
         Caption         =   "Fin"
         Height          =   240
         Left            =   345
         TabIndex        =   13
         Top             =   1140
         Width           =   645
      End
   End
   Begin VB.Frame FrameRep 
      Caption         =   "Clientes"
      Height          =   2010
      Left            =   240
      TabIndex        =   9
      Top             =   1005
      Width           =   3000
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         TabIndex        =   4
         Top             =   1530
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1215
         MaxLength       =   20
         TabIndex        =   3
         Top             =   1065
         Width           =   1455
      End
      Begin VB.OptionButton OpTodos 
         Caption         =   "Todos"
         Height          =   270
         Left            =   330
         TabIndex        =   1
         Top             =   300
         Width           =   1125
      End
      Begin VB.OptionButton OpRango 
         Caption         =   "Rango"
         Height          =   255
         Left            =   330
         TabIndex        =   2
         Top             =   630
         Width           =   1125
      End
      Begin VB.Label Label3 
         Caption         =   "Fin"
         Height          =   255
         Left            =   540
         TabIndex        =   11
         Top             =   1575
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Inicio "
         Height          =   255
         Left            =   540
         TabIndex        =   10
         Top             =   1035
         Width           =   735
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   735
      Left            =   2010
      Picture         =   "FrmGuias.frx":08F0
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3165
      Width           =   735
   End
   Begin VB.CommandButton Command7 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   735
      Left            =   3570
      Picture         =   "FrmGuias.frx":0D32
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3180
      Width           =   735
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   225
      Top             =   3075
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "FrmGuias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim db As Database
Dim almacen As String
Dim Conexion As String
Dim Adodc3 As ADODB.Recordset

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub Command7_Click()
  MousePointer = vbHourglass
'  If Frame1.Visible And Frame2.Visible Then
        MousePointer = vbDefault
        Unload Me
'  Else
'        Frame1.Visible = True
'        Frame2.Visible = True
'        FrameRep.Visible = False
'        MousePointer = vbDefault
'  End If
End Sub

Private Sub Command1_Click()
imprimir
End Sub

Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then DTPicker2.SetFocus
End Sub

Private Sub DTPicker2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Command1.SetFocus
End Sub

Private Sub Form_Load()
Frame1.Visible = True
Carga_Almacen
Combo1.ListIndex = 0
central FrmGuias
OpTodos.Value = True
DTPicker1 = DateAdd("m", -1, Date)
DTPicker2 = Date
End Sub

Private Sub OpRango_Click()
If OpRango.Value Then
    Text1.Enabled = True
    Text2.Enabled = True
    Text1.SetFocus
End If
End Sub

Private Sub Carga_Almacen()
Dim RSQL As String
Dim rs As Recordset
 
RSQL = "select  TADESCRI FROM TabAlm "
'Set db = Workspaces(0).OpenDatabase(cRuta2)
'Set RS = db.OpenRecordset(RSQL, dbOpenSnapshot)
Set rs = cConexCom.Execute(RSQL)
While Not rs.EOF
'    Combo1.AddItem (rS(0))
    rs.MoveNext
Wend
rs.Close
'Dim Esql As String
'Esql = "Select Campo from material"
'Set Rs = Db.OpenResultset(Esql, rdOpenKeyset)
'criterio = "CJA_CODIGO = " & Chr$(34) + Text1.Text + Chr$(34)
''Data1.Recordset.MoveFisrt
'Cmbgrupo.AddItem ("Todas los grupos")
'Rs.MoveFirst
'While Not Data1.Recordset.EOF
'Cmbgrupo.AddItem (Data1.Recordset.Columns(3))
' Data1.Recordset.MoveNext
''Wend
'Data1.Recordset.Close
End Sub

Private Sub imprimir()
Dim Codigo1 As String
Dim Codigo2 As String
Dim cadena As String
Dim cTip As String, Tipo As String
Codigo1 = Trim(Text1)
Screen.MousePointer = 11
CrystalReport1.WindowTitle = "Inv093 -- Control de Inventarios"
CrystalReport1.ReportFileName = cRutP & "inv093.rpt"

cadena = "({MOVALMCAB.CAFECDOC} IN DATE (" & Format(DTPicker1, "yyyy") & "," & Format(DTPicker1, "mm") & "," & Format(DTPicker1, "dd") & ") "
cadena = cadena & "to DATE (" & Format(DTPicker2, "yyyy") & "," & Format(DTPicker2, "mm") & "," & Format(DTPicker2, "dd") & ")) "
cadena = cadena & "and {MOVALMCAB.CATD}='GS' AND {MOVALMCAB.CAALMA}='" & VGAlma & "'"

If Combo1.ListIndex = 0 Then
  Tipo = "FACTURADAS"
  cTip = "F"
Else
  Tipo = "PENDIENTES DE PAGO"
  cTip = "P"
End If

If OpTodos.Value Then
    cadena = cadena & " And {MOVALMCAB.CASITGUI}='" & cTip & "'"
    Ubi_Tab CrystalReport1
    CrystalReport1.DiscardSavedData = True
    CrystalReport1.Destination = crptToWindow
    CrystalReport1.SelectionFormula = cadena
    CrystalReport1.WindowShowPrintBtn = True
    CrystalReport1.WindowShowRefreshBtn = True
    CrystalReport1.WindowShowSearchBtn = True
    CrystalReport1.WindowShowPrintSetupBtn = True
    CrystalReport1.Formulas(0) = "emp = '" & VGNemp & "'"
    CrystalReport1.Formulas(1) = "hora ='" & Format(Time, "hh:mm:ss") & "'"
    CrystalReport1.Formulas(2) = "Tipo ='" & Tipo & "'"
    CrystalReport1.Formulas(3) = "xalmacen ='" & VGNomAlm & "'"
   ' CrystalReport1.WindowTitle = "Reporte de Guías de Remisión " & Tipo VGNomAlm

    If CrystalReport1.Status <> 2 Then CrystalReport1.Action = 1
    Screen.MousePointer = 1
    Exit Sub
End If

If Text1 = "" Then
        MsgBox "Ingrese el codigo", vbExclamation, "Error"
        Text1.SetFocus
        Exit Sub
End If
If OpRango.Value Then           'Un select
    
    If Text2 <> "" Then
            Codigo2 = Text2         '  "23134671"
            cadena = cadena & " and ({MOVALMCAB.CACODCLI} in '" & Codigo1 & "' to '" & Codigo2 & "')  And {MOVALMCAB.CASITGUI}='" & cTip & "'"
    Else
            Codigo2 = Codigo1
            cadena = cadena & " and {MOVALMCAB.CACODCLI} = '" & Codigo1 & "' And {MOVALMCAB.CASITGUI}='" & cTip & "'"
    End If
    
    Ubi_Tab CrystalReport1
    CrystalReport1.DiscardSavedData = True
    CrystalReport1.Destination = crptToWindow
    CrystalReport1.SelectionFormula = cadena
    CrystalReport1.Formulas(0) = "emp = '" & VGNemp & "'"
    CrystalReport1.Formulas(1) = "hora ='" & Format(Time, "hh:mm:ss") & "'"
    CrystalReport1.Formulas(2) = "Tipo ='" & Tipo & "'"
   ' CrystalReport1.WindowTitle = "Reporte de Guías de Remisión " & Tipo
    If CrystalReport1.Status <> 2 Then CrystalReport1.Action = 1
End If
Screen.MousePointer = 1
End Sub

Private Sub OpTodos_Click()
Text1.Enabled = False
Text2.Enabled = False
limpiar_t1_t2
End Sub

Private Sub OpTodos_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then DTPicker1.SetFocus
End Sub

Private Sub Text1_DblClick()
Dim Adodc2 As ADODB.Recordset
Set Adodc2 = New ADODB.Recordset
         Adodc2.Open "Select CCODCLI, CNOMCLI from MaeCli", cConexCom, adOpenStatic, adLockOptimistic
         frmReferencia.Conectar Adodc2, "Select CCODCLI, CNOMCLI from MaeCli"
         frmReferencia.Label1.Caption = "Clientes"
         frmReferencia.Show vbModal
         Adodc2.Close
         If vGUtil(1) <> "" Then
                 Text1 = (vGUtil(1))
         End If
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

Private Sub Text1_KeyPress(KeyAscii As Integer)
If Not OpRango.Value Then
   OpRango = True
End If
If KeyAscii = 13 And Text1 <> "" Then
        If Existe(1, Text1, "MAECLI", "CCODCLI", False) = False Then
                MsgBox "El código de Cliente no existe", vbInformation, mensaje1
                Text1.SetFocus: Exit Sub
         Else
                Text2.Enabled = True
                Text2.SetFocus
         End If
 End If
End Sub

Private Sub Text2_DblClick()
Dim Adodc2 As ADODB.Recordset
Set Adodc2 = New ADODB.Recordset

  Adodc2.Open "Select CCODCLI, CNOMCLI from MaeCli", cConexCom, adOpenStatic, adLockOptimistic
  frmReferencia.Conectar Adodc2, "Select CCODCLI, CNOMCLI from MaeCli"
  frmReferencia.Label1.Caption = "Clientes"
  frmReferencia.Show vbModal
  Adodc2.Close

  If vGUtil(1) <> "" Then
      Text2 = (vGUtil(1))
  End If
  If Text2 <> "" Then
       DTPicker1.SetFocus
  End If
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then Text2_DblClick
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 And Text2 <> "" Then
     If OpRango.Value Then
        If Existe(1, Text2, "MAECLI", "CCODCLI", False) = False Then
             MsgBox "El código de Cliente no existe", vbInformation, mensaje1
             Text2.SetFocus: Exit Sub
        End If
        If Text1 > Text2 Then
               MsgBox "El codigo fin debe ser mayor que el inicio", vbInformation, mensaje1
               Exit Sub
        End If
        DTPicker1.SetFocus
      End If
  ElseIf KeyAscii = 13 And Text2 = "" Then
      Command1.SetFocus
  End If
End Sub

Function Existe_cod_art(text As TextBox) As String
Dim rs As Recordset
Dim RSQL As String
RSQL = "select  ACODIGO FROM maeart where ACODIGO = '" & text & "'" '
'Set db = Workspaces(0).OpenDatabase(cRuta2)
'Set RS = db.OpenRecordset(RSQL, dbOpenSnapshot)
Set rs = cConexCom.Execute(RSQL)
If Not rs.EOF Then
    Existe_cod_art = rs(0)
Else
    MsgBox "El tipo de codigo no existe !", vbOKOnly, "Error"
    Existe_cod_art = ""
End If
rs.Close
End Function

Private Sub limpiar_t1_t2()
Text1 = ""
Text2 = ""
DTPicker1.Value = Date
DTPicker2.Value = Date
End Sub
