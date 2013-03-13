VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form FrmArLineas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lineas de Articulos"
   ClientHeight    =   4725
   ClientLeft      =   1935
   ClientTop       =   2145
   ClientWidth     =   7335
   ControlBox      =   0   'False
   LinkTopic       =   "Form5"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4725
   ScaleMode       =   0  'User
   ScaleWidth      =   7407.985
   ShowInTaskbar   =   0   'False
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   600
      Top             =   2760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame5 
      Height          =   735
      Left            =   240
      TabIndex        =   10
      Top             =   0
      Width           =   6855
      Begin VB.TextBox TxFiltro 
         Height          =   300
         Left            =   1200
         MaxLength       =   45
         TabIndex        =   12
         Text            =   "TxFiltro"
         Top             =   240
         Width           =   2655
      End
      Begin VB.ComboBox CmbOrden 
         Height          =   315
         ItemData        =   "FrmArLinea.frx":0000
         Left            =   5040
         List            =   "FrmArLinea.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label32 
         Caption         =   "Buscar  :"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label33 
         Caption         =   "Orden :"
         Height          =   255
         Left            =   4320
         TabIndex        =   13
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   240
      TabIndex        =   1
      Top             =   3480
      Width           =   6855
      Begin VB.CommandButton command5 
         Caption         =   "&Reporte"
         Height          =   675
         Left            =   3600
         Picture         =   "FrmArLinea.frx":0027
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   240
         Width           =   775
      End
      Begin VB.CommandButton Command7 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         Height          =   675
         Left            =   5760
         Picture         =   "FrmArLinea.frx":0469
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   240
         Width           =   775
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Ingreso"
         Height          =   675
         Left            =   360
         Picture         =   "FrmArLinea.frx":08AB
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   240
         Width           =   775
      End
      Begin VB.CommandButton Command19 
         Caption         =   "&Grupos"
         Height          =   675
         Left            =   4680
         Picture         =   "FrmArLinea.frx":0CED
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   240
         Width           =   775
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Eliminar"
         Height          =   675
         Left            =   2520
         Picture         =   "FrmArLinea.frx":348F
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   240
         Width           =   775
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Modificar"
         Height          =   675
         Left            =   1440
         Picture         =   "FrmArLinea.frx":38D1
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   240
         Width           =   775
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   240
      TabIndex        =   7
      Top             =   3480
      Visible         =   0   'False
      Width           =   6855
      Begin VB.CommandButton Command9 
         Caption         =   "&Cancelar"
         Height          =   675
         Left            =   3960
         Picture         =   "FrmArLinea.frx":3D13
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   775
      End
      Begin VB.CommandButton Command8 
         Caption         =   "&Grabar"
         Height          =   675
         Left            =   1800
         Picture         =   "FrmArLinea.frx":4155
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   775
      End
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "FrmArLinea.frx":4597
      Height          =   2610
      Left            =   240
      OleObjectBlob   =   "FrmArLinea.frx":45AB
      TabIndex        =   0
      Top             =   840
      Width           =   6855
   End
   Begin VB.Frame Frame3 
      Height          =   2895
      Left            =   600
      TabIndex        =   2
      Top             =   360
      Visible         =   0   'False
      Width           =   6135
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   2280
         MaxLength       =   6
         TabIndex        =   4
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   2280
         MaxLength       =   45
         TabIndex        =   3
         Top             =   1680
         Width           =   3495
      End
      Begin VB.Label Label2 
         Caption         =   "Código:"
         Height          =   255
         Left            =   720
         TabIndex        =   6
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Nombre:"
         Height          =   255
         Left            =   720
         TabIndex        =   5
         Top             =   1680
         Width           =   855
      End
   End
End
Attribute VB_Name = "FrmArLineas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim resp As String
Public Fam As String
Dim nTra As Integer
Dim Data1 As New ADODB.Recordset
Dim Data2 As New ADODB.Recordset

Private Sub command5_Click()
Dim CTIME As String
If Data1.RecordCount > 0 Then

CTIME = Format(Time, "hh:mm:ss")

With FrmArFam
    .CrystalReport1.WindowTitle = "Inv039  -- Control de Inventarios"
    .CrystalReport1.ReportFileName = cRutP & "inv039.Rpt"
    Call Ubi_Tab(.CrystalReport1)
    .CrystalReport1.WindowShowPrintBtn = True
    .CrystalReport1.WindowShowRefreshBtn = True
    .CrystalReport1.WindowShowSearchBtn = True
    .CrystalReport1.WindowShowPrintSetupBtn = True
    .CrystalReport1.Formulas(0) = "Hora = '" & CTIME & "'"
    .CrystalReport1.Formulas(1) = "Empresa = '" & Mid(VGNemp, 1, 20) & "'"
    .CrystalReport1.Formulas(2) = "Familia = '" & Mid(FrmArFam.Data1.Recordset("FAM_NOMBRE"), 1, 17) & "'"
    .CrystalReport1.Formulas(3) = ""
    .CrystalReport1.SelectionFormula = "{LINEAS.FAM_CODIGO} =  '" & Trim(Fam) & "'"
    .CrystalReport1.WindowTop = 100
    .CrystalReport1.WindowLeft = 150
    .CrystalReport1.DiscardSavedData = True
    .CrystalReport1.Destination = crptToWindow
    If .CrystalReport1.Status <> 2 Then .CrystalReport1.Action = 1
End With
End If
End Sub

Private Sub Command9_Click()
DBGrid1.Visible = True
Command19.Visible = True
Frame5.Visible = True
Frame2.Visible = True
Frame1.Visible = False
Frame3.Visible = False
DBGrid1.SetFocus
End Sub

Private Sub TxFiltro_Change()
If Data1.RecordCount > 0 Then
    If Trim(TxFiltro) <> "" Then
        'Data1.Recordset.MoveFirst
        If CmbOrden.ListIndex = 0 Then
           ' Data1.Recordset.FindFirst "LIN_CODIGO like '" & Trim(UCase(TxFiltro)) & "*'"
        ElseIf CmbOrden.ListIndex = 1 Then
          '  Data1.Recordset.FindFirst "LIN_NOMBRE like '" & Trim(UCase(TxFiltro)) & "*'"
        End If
        'If Data1.Recordset.EOF Then Data1.Recordset.MoveFirst
    End If
End If
End Sub

Private Sub CmbOrden_Click()             ' Ordenar por
'Dim nCom As Integer
'
'nCom = CmbOrden.ListIndex
'
'Select Case nCom
'Case 0
'    Data1.RecordSource = "Select * from LINEAS where FAM_CODIGO='" & FrmArFam.Data1.Recordset("FAM_CODIGO") & "' order by LIN_CODIGO"
'Case 1
'    Data1.RecordSource = "Select * from LINEAS where FAM_CODIGO='" & FrmArFam.Data1.Recordset("FAM_CODIGO") & "' order by LIN_NOMBRE"
'End Select
'TxFiltro = ""
'Data1.Refresh
'If DBGrid1.Visible Then DBGrid1.SetFocus
End Sub



'Sub listado(wcad)
'  Set DBGrid1.DataSource = Nothing
'  Set RS = Nothing
'
'  Set RS = cConexCom.Execute(wcad)
'  Set DBGrid1.DataSource = RS
'  With DBGrid1
'      .Columns(0).Caption = "Codigo"
'      .Columns(0).Width = 1000
'      .Columns(1).Caption = "Descripcion"
'      .Columns(1).Width = 3800
'      .Columns(2).Caption = "Cuenta Contable"
'      .Columns(2).Width = 1000
'      .MarqueeStyle = dbgHighlightRow
'      .Refresh
'  End With
'
'End Sub


Private Sub DataGrid1_KeyPress(KeyAscii As Integer)

If KeyAscii = vbKeyBack Then
    If Len(TxFiltro) - 1 > 0 Then
        TxFiltro = Left(TxFiltro, Len(TxFiltro) - 1)
    Else
        TxFiltro = ""
    End If
    KeyAscii = 0
ElseIf KeyAscii <> 13 Then
    TxFiltro = TxFiltro & Chr(KeyAscii)
End If
End Sub

Private Sub Command1_Click()
resp = "S"
Limpiar
Text1.Enabled = True
DBGrid1.Visible = False
Frame2.Visible = False
Frame5.Visible = False
Frame3.Caption = "Ingreso de Lineas"

Frame1.Visible = True
Frame3.Visible = True
Text1.SetFocus
End Sub

Private Sub Command19_Click()
'If Data1.Recordset.RecordCount > 0 Then FrmArGrupos.Show 1
End Sub

Private Sub Command2_Click()
'If Data1.Recordset.RecordCount > 0 Then
'    Limpiar
'    resp = "N"
'    Frame3.Caption = "Modificación de Lineas"
'
'    DBGrid1.Visible = False
'    Command19.Visible = False
'    Frame2.Visible = False
'    Frame5.Visible = False
'    Frame1.Visible = True
'    Frame3.Visible = True
'
'    Text1.text = Data1("LIN_CODIGO")
'    Text1.Enabled = False
'
'    If Not IsNull(Data1("LIN_NOMBRE")) Then
'      Text2.text = Data1("LIN_NOMBRE")
'    Else
'      Text2.text = ""
'    End If
'   Text2.SetFocus
'End If
End Sub

Private Sub Command3_Click()
'On Error GoTo EliErr
Dim cSql1 As String
Dim CSQL2 As String, cSql3 As String
Dim cCodigo1 As String
Dim cSel1 As Recordset
Dim cCodigo As String

'If Data1.Recordset.RecordCount > 0 Then
'    CSQL2 = "Delete from GRUPO Where FAM_CODIGO='" & FrmArFam.Data1.Recordset("FAM_CODIGO") & "' AND LIN_CODIGO= '" & Data1.Recordset("LIN_CODIGO") & "'"
'
'    Dim cSqlA As String, cSelA As ADODB.Recordset
'
'    cSqlA = "Select * FROM GRUPO WHERE FAM_CODIGO='" & FrmArFam.Data1.Recordset("FAM_CODIGO") & "' AND LIN_CODIGO = '" & Trim(Data1.Recordset("LIN_CODIGO")) & "'"
'    Set cSelA = New ADODB.Recordset
'    cSelA.Open cSqlA, cConexCom, adOpenStatic
'    If cSelA.RecordCount > 0 Then
'        If MsgBox("La Linea seleccionada tiene registrado Grupos, al Eliminarla eliminará sus Grupos, desea Eliminarla de todas maneras", vbYesNo, "Eliminacion de Registro") = vbNo Then
'            cSelA.Close: Exit Sub
'       End If
'    End If
'    cSelA.Close
'
'
'    If MsgBox("Seguro de Eliminar", vbQuestion + vbOKCancel, mensaje1) = vbOK Then
'            nTra = 2
'            cCodigo1 = Pos_Dato1(Data1.Recordset, "LIN_CODIGO")
'            nTra = 1
'            cConexCom.BeginTrans
'            cConexCom.Execute CSQL2
'            cConexCom.CommitTrans
'            nTra = 0
'
'            Data1.Refresh
'
'            If cCodigo1 <> "" Then
'                Data1.Recordset.FindFirst "LIN_CODIGO='" & cCodigo1 & "'"
'            End If
'    End If
'    DBGrid1.Refresh
'    If DBGrid1.Visible Then DBGrid1.SetFocus
'Else
'    MsgBox "No existe ningún registro para Eilminar", vbInformation, mensaje1
'End If
'Exit Sub
'EliErr:
'    MsgBox Err.Description
'    If nTra = 1 Then cConexCom.RollbackTrans
End Sub

Private Sub Command7_Click()
Unload Me
 End Sub

Private Sub Command8_Click()
On Error GoTo GrabErr
Dim cFam As String

'   resp = "S"
If resp = "S" Then
  If Text1 = "" Then
     MsgBox "Ingrese Código de Linea ", vbInformation, "Mensaje"
     Text1.SetFocus
     Exit Sub
  Else
     If Existe(1, Trim(Text1), "LINEAS", "LIN_CODIGO", False, Fam, "FAM_CODIGO") Then
          MsgBox "El código de Linea ya existe", vbInformation, "Mensaje"
          Text1.SetFocus
          Exit Sub
       End If
  End If
End If

If Text2 = "" Then
   MsgBox "Ingrese Descripción de Linea", vbExclamation, "Aviso"
   Text2.SetFocus
   Exit Sub
End If
  
If resp = "S" Then
    Data1.Recordset.AddNew
Else
    'Data1.Edit
End If
Data1("FAM_CODIGO") = Fam
Data1("LIN_CODIGO") = Text1.text
If Not IsNull(Text2.text) Then
    Data1("LIN_NOMBRE") = Text2.text
Else
    Data1("LIN_NOMBRE") = " "
End If
Data1.Update
Data1.Refresh
DBGrid1.Refresh
   
Data1.Find "LIN_CODIGO = '" & Text1.text & "'"
   
If resp = "S" Then
      Limpiar
      Text1.SetFocus
Else
      'Label1.Visible = True
      DBGrid1.Visible = True
      Command19.Visible = True
      Frame5.Visible = True
      Frame2.Visible = True
      Frame1.Visible = False
      Frame3.Visible = False
      DBGrid1.SetFocus
End If
Exit Sub
GrabErr:
    MsgBox Err.Description
    'If nTra = 1 Then cConexCom.RollbackTrans
End Sub

Sub Limpiar()
Text1 = ""
Text2 = ""
End Sub

Private Sub Form_Activate()
TxFiltro = ""
CmbOrden.ListIndex = 0
If DBGrid1.Visible Then DBGrid1.SetFocus
End Sub

Private Sub Form_Load()
central Me
Init_ControlDBGrid DBGrid1
Me.Caption = "Lineas de la Familia :  " & Mid(FrmArFam.Data1.Recordset("FAM_NOMBRE"), 1, 20)
'Data1.DatabaseName = cRuta2
'Data1.RecordSource = "Select * from LINEAS where FAM_CODIGO='" & FrmArFam.Data1.Recordset("FAM_CODIGO") & "' order by LIN_CODIGO"
Data1.Open "Select * from LINEAS where FAM_CODIGO='" & FrmArFam.Data1.Recordset("FAM_CODIGO") & "' order by LIN_CODIGO", cConexCom, adOpenDynamic, adLockOptimistic

'Data1.Refresh
Fam = FrmArFam.Data1.Recordset("FAM_CODIGO")

'Command19.Visible = True
End Sub

Private Sub Text1_GotFocus()
Enfoque Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
Dim cFam As String

If KeyAscii = 13 Then
    If Trim(Text1) <> "" Then
       If Existe(1, Trim(Text1), "LINEAS", "LIN_CODIGO", False, Fam, "FAM_CODIGO") Then
          MsgBox "El código de Linea ya existe", vbInformation, "Mensaje"
          Text1 = "": Text1.SetFocus
          Exit Sub
       End If
    Else
          MsgBox "Ingrese código de Linea", vbInformation, "Mensaje"
          Text1 = "": Text1.SetFocus
    End If
    Text2.SetFocus
Else
    If KeyAscii = 39 Or KeyAscii = 34 Then KeyAscii = 0
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End If
End Sub
Private Sub Text2_GotFocus()
Enfoque Text2
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Trim(Text2) = "" Then
       MsgBox "Ingrese Descripcion de Lineaa", vbInformation, "Mensaje"
       Text2 = "": Text2.SetFocus
    End If
    Command8.SetFocus
End If
End Sub
