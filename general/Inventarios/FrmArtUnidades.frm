VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form FrmArUnidades 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Unidades de Medida"
   ClientHeight    =   4170
   ClientLeft      =   1770
   ClientTop       =   1710
   ClientWidth     =   7170
   ControlBox      =   0   'False
   LinkTopic       =   "Form5"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleMode       =   0  'User
   ScaleWidth      =   7241.343
   ShowInTaskbar   =   0   'False
   Begin MSDataGridLib.DataGrid DbGrid1 
      Height          =   2025
      Left            =   180
      TabIndex        =   20
      Top             =   840
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   3572
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   6720
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame3 
      Height          =   2415
      Left            =   420
      TabIndex        =   1
      Top             =   600
      Visible         =   0   'False
      Width           =   6135
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   2265
         MaxLength       =   5
         TabIndex        =   3
         Top             =   675
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   2280
         MaxLength       =   20
         TabIndex        =   2
         Top             =   1395
         Width           =   3495
      End
      Begin VB.Label Label2 
         Caption         =   "Código             :"
         Height          =   255
         Left            =   705
         TabIndex        =   5
         Top             =   675
         Width           =   1440
      End
      Begin VB.Label Label3 
         Caption         =   "Descripción     :"
         Height          =   255
         Left            =   705
         TabIndex        =   4
         Top             =   1395
         Width           =   1335
      End
   End
   Begin VB.Frame Frame5 
      Height          =   735
      Left            =   135
      TabIndex        =   9
      Top             =   0
      Width           =   6855
      Begin VB.TextBox TxFiltro 
         Height          =   300
         Left            =   1200
         MaxLength       =   45
         TabIndex        =   11
         Text            =   "TxFiltro"
         Top             =   240
         Width           =   2655
      End
      Begin VB.ComboBox CmbOrden 
         Height          =   315
         ItemData        =   "FrmArtUnidades.frx":0000
         Left            =   5040
         List            =   "FrmArtUnidades.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label32 
         Caption         =   "Buscar  :"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label33 
         Caption         =   "Orden :"
         Height          =   255
         Left            =   4320
         TabIndex        =   12
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   180
      TabIndex        =   0
      Top             =   2895
      Width           =   6855
      Begin VB.CommandButton command5 
         Caption         =   "&Reporte"
         Height          =   675
         Left            =   3600
         Picture         =   "FrmArtUnidades.frx":0027
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   240
         Width           =   775
      End
      Begin VB.CommandButton Command7 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         Height          =   675
         Left            =   5760
         Picture         =   "FrmArtUnidades.frx":0469
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   255
         Width           =   775
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Eliminar"
         Height          =   675
         Left            =   2520
         Picture         =   "FrmArtUnidades.frx":08AB
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   240
         Width           =   775
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Ingreso"
         Height          =   675
         Left            =   360
         Picture         =   "FrmArtUnidades.frx":0CED
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   240
         Width           =   775
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Modificar"
         Height          =   675
         Left            =   1440
         Picture         =   "FrmArtUnidades.frx":112F
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   240
         Width           =   775
      End
      Begin VB.CommandButton Command19 
         Caption         =   "&Equival."
         Height          =   675
         Left            =   4680
         Picture         =   "FrmArtUnidades.frx":1571
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   240
         Width           =   775
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1065
      Left            =   180
      TabIndex        =   6
      Top             =   2910
      Visible         =   0   'False
      Width           =   6840
      Begin VB.CommandButton Command9 
         Caption         =   "&Cancelar"
         Height          =   675
         Left            =   3960
         Picture         =   "FrmArtUnidades.frx":3D13
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   775
      End
      Begin VB.CommandButton Command8 
         Caption         =   "&Grabar"
         Height          =   675
         Left            =   1800
         Picture         =   "FrmArtUnidades.frx":4155
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   775
      End
   End
End
Attribute VB_Name = "FrmArUnidades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim resp As String
Dim nTra As Integer
Dim rs As New ADODB.Recordset

Private Sub command5_Click()
Dim CADENA As String
Dim cNomRepor  As String

cNomRepor = "unimedida.RPT"
If Trim(cNomRepor) <> "" Then
    CrystalReport1.Reset
    CrystalReport1.WindowTitle = "Reporte de Unidades de Medida"
    CrystalReport1.ReportFileName = VGParamSistem.RutaReport + "\" + cNomRepor
    CrystalReport1.Connect = VGcadenareport2
    CrystalReport1.StoredProcParam(0) = VGParamSistem.BDEmpresa
    
    CrystalReport1.formulas(0) = "emp ='" & VGparametros.RucEmpresa & "'"
    
    CrystalReport1.Destination = crptToWindow
    CrystalReport1.WindowShowPrintBtn = True
    CrystalReport1.WindowShowRefreshBtn = True
    CrystalReport1.WindowShowSearchBtn = True
    CrystalReport1.WindowShowPrintSetupBtn = True
    CrystalReport1.WindowState = crptMaximized
    If CrystalReport1.Status <> 2 Then
        CrystalReport1.Action = 1
    End If
Else
    MsgBox "No existe el nombre del Reporte, verifique en Formatos", vbInformation, "Información"
    Exit Sub
End If
End Sub

Private Sub Command9_Click()
DBGrid1.Visible = True
Command19.Visible = True
Frame5.Visible = True
Frame2.Visible = True
Frame1.Visible = False
Frame3.Visible = False
Set rs = VGCNx.Execute("select * from tabunimed")
DBGrid1.SetFocus
End Sub

Private Sub TxFiltro_Change()
Dim ncondi As String
'If RS.RecordCount > 0 Then
  If Trim(TxFiltro) <> "" Then
    If CmbOrden.ListIndex = 0 Then
         ncondi = Trim(UCase(TxFiltro)) & "%"
    ElseIf CmbOrden.ListIndex = 1 Then
         ncondi = Trim(UCase(TxFiltro)) & "%"
    End If
  Else
     ncondi = "SELECT * FROM TABUNIMED"
  End If
  Call Listado(ncondi)
'End If
End Sub

Private Sub CmbOrden_Click()             ' Ordenar por
Dim nCom As Integer
Dim nsql As String

nCom = CmbOrden.ListIndex

Select Case nCom
Case 0
    'Data1.RecordSource = "Select * from TABUNIMED  order by UM_ABREV"
    nsql = "Select * from TABUNIMED  order by UM_ABREV"
    Call Listado(nsql)
Case 1
    'Data1.RecordSource = "Select * from TABUNIMED  order by UM_NOMBRE"
    nsql = "Select * from TABUNIMED  order by UM_NOMBRE"
    Call Listado(nsql)
End Select
TxFiltro = ""
'Data1.Refresh
'If DbGrid1.Visible Then DbGrid1.SetFocus

End Sub

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
Command19.Visible = False
DBGrid1.Visible = False
Frame2.Visible = False
Frame5.Visible = False
Frame3.Caption = "Ingreso de Unidades de Medida"

Frame1.Visible = True
Frame3.Visible = True
Text1.SetFocus
End Sub

Private Sub Command19_Click()
If rs.RecordCount > 0 Then
    FrmArEquival.bdato = DBGrid1.Columns(0).text
    FrmArEquival.bdata = DBGrid1.Columns(1).text
    FrmArEquival.Show 1
End If
End Sub

Private Sub Command2_Click()
If rs.RecordCount > 0 Then
    Limpiar
    resp = "N"
    Frame3.Caption = "Modificación de Unidades de Medida"
    DBGrid1.Visible = False
    Command19.Visible = False
    Frame2.Visible = False
    Frame5.Visible = False
    Frame1.Visible = True
    Frame3.Visible = True

    Text1.text = rs.Fields("UM_ABREV")
    Text1.Enabled = False
    
    If Not IsNull(rs.Fields("UM_NOMBRE")) Then
          Text2.text = rs.Fields("UM_NOMBRE")
    Else
          Text2.text = ""
    End If
   Text2.SetFocus
End If
End Sub

Private Sub Command3_Click()
On Error GoTo EliErr
Dim cSql1 As String
Dim CSQL2 As String, cSql3 As String
Dim cCodigo1 As String
Dim cSel1 As Recordset
Dim cCodigo As String
Dim cSqlA As String
Dim cSelA As ADODB.Recordset
If rs.RecordCount > 0 Then
    
    cSqlA = "Select * FROM MovAlmDet WHERE DEUNIDAD = '" & Trim(DBGrid1.Columns(0)) & "'"
    Set cSelA = New ADODB.Recordset
    cSelA.Open cSqlA, VGCNx, adOpenStatic
    If Not cSelA.EOF Then
       MsgBox "Este Archivo Tiene enlace con Documentos no puede eliminarlo": cSelA.Close: Exit Sub
    End If
    cSelA.Close
    
    cSqlA = "Select * FROM FacDet WHERE DFUNIDAD = '" & Trim(rs.Fields("UM_ABREV")) & "'"
    cSelA.Open cSqlA, VGCNx, adOpenStatic
    If Not cSelA.EOF Then
       MsgBox "Este Archivo Tiene enlace con Documentos no puede eliminarlo": cSelA.Close: Exit Sub
    End If
    cSelA.Close
    
    
    cSql1 = "Delete from TABUNIMED  Where UM_ABREV= '" & rs.Fields("UM_ABREV") & "'"
 
    cSqlA = "Select * FROM TABEQUI WHERE EQUNIPRI = '" & Trim(rs.Fields("UM_ABREV")) & "'"
    Set cSelA = New ADODB.Recordset
    cSelA.Open cSqlA, VGCNx, adOpenStatic
    If cSelA.RecordCount > 0 Then
       If MsgBox("La Unidad de Medida seleccionada tiene registrada Unidades Equivalentes, al Eliminarla eliminará sus Unidades Equivalentes, desea Eliminarla de todas maneras", vbYesNo, "Eliminacion de Registro") = vbNo Then
            cSelA.Close: Exit Sub
       End If
    End If
    cSelA.Close
    

    If MsgBox("Seguro de Eliminar", vbQuestion + vbOKCancel, mensaje1) = vbOK Then
            nTra = 2
            'cCodigo1 = Pos_Dato1(RS, "UM_ABREV")
            nTra = 1
            VGCNx.BeginTrans
            VGCNx.Execute cSql1
            VGCNx.CommitTrans
            nTra = 0
            Set rs = VGCNx.Execute("Select * from TABUNIMED  order by UM_ABREV")
            Call Listado("")
    End If
    DBGrid1.Refresh
    
    If DBGrid1.Visible Then DBGrid1.SetFocus
Else
    MsgBox "No existe ningún registro para Eilminar", vbInformation, mensaje1
End If
Exit Sub
EliErr:
    MsgBox Err.Description
    If nTra = 1 Then VGCNx.RollbackTrans
End Sub

Private Sub Command7_Click()
   Unload Me
 End Sub

Private Sub Command8_Click()
On Error GoTo GrabErr
Dim cUni As String

If resp = "S" Then
  If Text1 = "" Then
     MsgBox "Ingrese Código de Unidad de Medida", vbInformation, "Mensaje"
     Text1.SetFocus
     Exit Sub
  Else
       If Existe(1, Trim(Text1), "TABUNIMED", "UM_ABREV", False) Then
            MsgBox "El código de Unidad de Medida ya existe", vbInformation, "Mensaje"
            Text1.SetFocus
            Exit Sub
       End If
  End If
End If

  If Text2 = "" Then
        MsgBox "Ingrese Descripción de Unidad de Medida", vbExclamation, "Aviso"
        Text2.SetFocus
        Exit Sub
  End If

    If resp = "S" Then
         VGCNx.Execute "INSERT INTO TABUNIMED " & _
                           "(UM_ABREV,UM_NOMBRE,UM_ESTADO)" & _
                           " VALUES (" & _
                           "'" & Text1.text & "'," & _
                           "'" & Text2.text & "'," & _
                           "'A')"
        Call Listado("SELECT * FROM TABUNIMED")
        
        Limpiar
        Text1.SetFocus
    Else
       VGCNx.Execute "UPDATE TABUNIMED " & _
                         " SET UM_NOMBRE ='" & Text2.text & "'," & _
                         " UM_ESTADO='A'" & _
                         " WHERE UM_ABREV ='" & Text1.text & "'"
    
        Call Listado("SELECT * FROM TABUNIMED")
        
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
    'If nTra = 1 Then Vgcnx.RollbackTrans
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
Set rs = VGCNx.Execute("Select * from TABUNIMED  order by UM_ABREV")
Call Listado("")
Command19.Visible = True
End Sub


Sub Listado(wcad)
  Set DBGrid1.DataSource = Nothing
  Dim nCursor As String
'  Set rs = VGcnx.Execute(wcad)
  If Trim(TxFiltro) <> "" And TxFiltro.Visible Then
         Select Case CmbOrden.ListIndex
        Case 0
            Set rs = VGCNx.Execute("SELECT * FROM TABUNIMED WHERE UM_ABREV LIKE '%" & wcad & "' ORDER BY 1")
        Case 1
            Set rs = VGCNx.Execute("SELECT * FROM TABUNIMED WHERE UM_NOMBRE LIKE '%" & wcad & "' ORDER BY 2 ")
        End Select
        If rs.EOF Then Set rs = VGCNx.Execute("SELECT * FROM TABUNIMED ")
 Else
        Select Case CmbOrden.ListIndex
        Case 0
            Set rs = VGCNx.Execute("SELECT * FROM TABUNIMED ORDER BY 1")
        Case 1
            Set rs = VGCNx.Execute("SELECT * FROM TABUNIMED ORDER BY 2 ")
        End Select
 End If
 Set DBGrid1.DataSource = rs
With DBGrid1
      .Columns(0).Caption = "Codigo"
      .Columns(0).Width = 1000
      .Columns(1).Caption = "Descripcion"
      .Columns(1).Width = 3800
      .MarqueeStyle = dbgHighlightRow
      .Refresh
End With
End Sub


Private Sub Text1_GotFocus()
   Enfoque Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
Dim cFam As String

If KeyAscii = 13 Then
    If Trim(Text1) <> "" Then
       If Existe(1, Trim(Text1), "TABUNIMED", "UM_ABREV", False) Then
          MsgBox "El código de Unidad de Medida ya existe", vbInformation, "Mensaje"
          Text1 = "": Text1.SetFocus
          Exit Sub
       End If
    Else
          MsgBox "Ingrese código de Unidad de Medida", vbInformation, "Mensaje"
          Text1 = "": Text1.SetFocus
    End If
    Text2.SetFocus
Else
    'If KeyAscii = 39 Or KeyAscii = 34 Then KeyAscii = 0
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End If
End Sub
Private Sub Text2_GotFocus()
Enfoque Text2
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Trim(Text2) = "" Then
       MsgBox "Ingrese Descripcion de Unidad de Medida", vbInformation, "Mensaje"
       Text2 = "": Text2.SetFocus
    End If
    Command8.SetFocus
End If
End Sub
