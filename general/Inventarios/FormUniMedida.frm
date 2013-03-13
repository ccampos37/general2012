VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form FrmArUniMed 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Unidades"
   ClientHeight    =   4740
   ClientLeft      =   1680
   ClientTop       =   1080
   ClientWidth     =   7980
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4740
   ScaleWidth      =   7980
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "FormUniMedida.frx":0000
      Height          =   2415
      Left            =   240
      OleObjectBlob   =   "FormUniMedida.frx":0014
      TabIndex        =   0
      Top             =   840
      Width           =   7455
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   30
      Top             =   3765
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame4 
      Height          =   1095
      Left            =   600
      TabIndex        =   16
      Top             =   3360
      Width           =   6735
      Begin VB.CommandButton Command2 
         Caption         =   "&Modificar"
         Height          =   675
         Left            =   2220
         Picture         =   "FormUniMedida.frx":0A0F
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   240
         Width           =   775
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Ingreso"
         Height          =   675
         Left            =   1110
         Picture         =   "FormUniMedida.frx":0E51
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   240
         Width           =   775
      End
      Begin VB.CommandButton Command7 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         Height          =   675
         Left            =   4440
         Picture         =   "FormUniMedida.frx":1293
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   240
         Width           =   775
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Eliminar"
         Height          =   675
         Left            =   3330
         Picture         =   "FormUniMedida.frx":16D5
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   240
         Width           =   775
      End
   End
   Begin VB.Data Data2 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Inventarios.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   615
      Left            =   5865
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "TabEqui"
      Top             =   2520
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   615
      Left            =   5640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   840
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Frame Frame5 
      Height          =   735
      Left            =   240
      TabIndex        =   18
      Top             =   0
      Width           =   7455
      Begin VB.ComboBox CmbOrden 
         Height          =   315
         ItemData        =   "FormUniMedida.frx":1B17
         Left            =   5520
         List            =   "FormUniMedida.frx":1B21
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox TxFiltro 
         Height          =   300
         Left            =   1200
         MaxLength       =   20
         TabIndex        =   19
         Text            =   "TxFiltro"
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label Label33 
         Caption         =   "Orden :"
         Height          =   255
         Left            =   4560
         TabIndex        =   22
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label32 
         Caption         =   "Buscar  :"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3210
      Left            =   555
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   6735
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   2520
         MaxLength       =   20
         TabIndex        =   5
         Top             =   1680
         Width           =   3615
      End
      Begin VB.CommandButton Command19 
         Caption         =   "&Equivalencias"
         Height          =   375
         Left            =   2520
         TabIndex        =   6
         Top             =   2280
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   2520
         MaxLength       =   6
         TabIndex        =   3
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Nombre :"
         Height          =   255
         Left            =   1080
         TabIndex        =   4
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Codigo :"
         Height          =   255
         Left            =   1080
         TabIndex        =   2
         Top             =   960
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1020
      Left            =   600
      TabIndex        =   7
      Top             =   3420
      Visible         =   0   'False
      Width           =   6735
      Begin VB.CommandButton Command20 
         Caption         =   "&Grabar"
         Height          =   675
         Left            =   2280
         Picture         =   "FormUniMedida.frx":1B3E
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   240
         Width           =   775
      End
      Begin VB.CommandButton Command21 
         Caption         =   "&Salir"
         Height          =   675
         Left            =   3480
         Picture         =   "FormUniMedida.frx":1F80
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   775
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Relacion de Equivalencias"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   3015
      Left            =   600
      TabIndex        =   9
      Top             =   300
      Visible         =   0   'False
      Width           =   6735
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   2505
         MaxLength       =   6
         TabIndex        =   15
         Top             =   1335
         Width           =   1215
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   2505
         TabIndex        =   11
         Top             =   1785
         Width           =   1695
      End
      Begin VB.Label Label7 
         Caption         =   "Label7"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3960
         TabIndex        =   17
         Top             =   1320
         Width           =   2535
      End
      Begin VB.Label Label6 
         Caption         =   "Unidad Equivalente :"
         Height          =   255
         Left            =   465
         TabIndex        =   14
         Top             =   1365
         Width           =   1815
      End
      Begin VB.Label Label5 
         Caption         =   "Factor :"
         Height          =   255
         Left            =   465
         TabIndex        =   13
         Top             =   1815
         Width           =   1455
      End
      Begin VB.Label Label4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   12
         Top             =   840
         Width           =   3735
      End
      Begin VB.Label Label3 
         Caption         =   "Unidad Principal :"
         Height          =   255
         Left            =   465
         TabIndex        =   10
         Top             =   960
         Width           =   1455
      End
   End
End
Attribute VB_Name = "FrmArUniMed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nOpc As Byte
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

Private Sub Form_Activate()
TxFiltro = ""
CmbOrden.ListIndex = 0
If DBGrid1.Visible Then DBGrid1.SetFocus
End Sub

Private Sub TxFiltro_Change()
If Data1.Recordset.RecordCount > 0 Then
    If Trim(TxFiltro) <> "" Then
        Data1.Recordset.MoveFirst
        
        If CmbOrden.ListIndex = 0 Then
            Data1.Recordset.FindFirst "UM_ABREV like '" & Trim(UCase(TxFiltro)) & "*'"
        ElseIf CmbOrden.ListIndex = 1 Then
            Data1.Recordset.FindFirst "UM_NOMBRE like '" & Trim(UCase(TxFiltro)) & "*'"
        End If
        If Data1.Recordset.EOF Then Data1.Recordset.MoveFirst
        
    End If
End If
End Sub

Private Sub CmbOrden_Click()             ' Ordenar por
Dim nCom As Integer

nCom = CmbOrden.ListIndex

Select Case nCom
Case 0
    Data1.RecordSource = "Select * from TABUNIMED order by UM_ABREV"
    
Case 1
    Data1.RecordSource = "Select * from TABUNIMED order by UM_NOMBRE"
End Select
'TxFiltro = ""
Data1.Refresh
If DBGrid1.Visible Then DBGrid1.SetFocus
End Sub


Private Sub Command1_Click()
  nOpc = 0
  Frame5.Visible = False
  Frame1.Visible = True
  Frame2.Visible = True
  Frame1.Caption = "Ingreso de Unidad de Medida"
  Limpiar
   '@@@@@@@@@@@
  DBGrid1.Visible = False
  Frame4.Visible = False
  Text1.Enabled = True
  Text1.SetFocus
   '@@@@@@@@@@@
End Sub

Private Sub Command19_Click()
If Text1 <> "" And Text2 <> "" Then
   Label4 = Mid(Text2, 1, 20)
   Frame2.Visible = True
   Frame3.Visible = True
   Frame1.Visible = False
   Text4.SetFocus
Else
 If Text1 = "" And nOpc = 0 Then
 MsgBox "Registre Codigo de Unidad", vbInformation, "Aviso"
 Text1.SetFocus
 Else
 MsgBox "Registre Descripcion de Unidad", vbInformation, "Aviso"
 Text2.SetFocus
 End If
End If
End Sub

Private Sub Command2_Click()
If Data1.Recordset.RecordCount > 0 Then
    If Data1.Recordset("UM_ESTADO") = "B" Then
      MsgBox "No se puede modificar dicha unidad", vbInformation, "Aviso"
      DBGrid1.SetFocus: Exit Sub
    End If
    nOpc = 1
    Frame1.Visible = True
    Frame1.Caption = "Modificacion de Unidad de Medida"
    
    Frame5.Visible = False
    Frame2.Visible = True
    DBGrid1.Visible = False
    Frame4.Visible = False
    Limpiar
    Text1 = Data1.Recordset("UM_ABREV")
    
    If IsNull(Data1.Recordset("UM_NOMBRE")) Then
    Text2.text = ""
    Else
    Text2.text = Data1.Recordset("UM_NOMBRE")
    End If
    Text1.Enabled = False
    Text2.SetFocus
 End If
End Sub

Private Sub Command20_Click()
On Error GoTo GrabErr
Dim cUni As String
If nOpc = 0 Then  'Ingreso
   If Trim(Text1) <> "" Then
         If Existe(1, Text1, "TABUNIMED", "UM_ABREV", False) Then
            MsgBox "Unidad de Medida ya esta registrada", vbInformation, "Mensaje"
            If Frame3.Visible = True Then
               Frame3.Visible = False
               Frame1.Visible = True
               Text3 = ""
               Text4 = ""
               Label7 = ""
             Else
               Frame1.Visible = True
             End If
            Text1.SetFocus: Exit Sub
        End If
   Else
        MsgBox "Ingrese Codigo de Unidad de Medida", vbInformation, "Mensaje"
        Text1.SetFocus: Exit Sub
   End If
End If

If Frame3.Visible Then   'Señalo Unidades Equivalentes
 
   If Text3 = "" Then Text3 = "0"
   
    If Trim(Text4) <> "" Then
        If Trim(Label7) = "" Then
            MsgBox "Unidad de Medida no registrada", vbInformation, "Mensaje"
            Text4.SetFocus: Exit Sub
        End If
    Else
        MsgBox "Ingrese Unidad de Medida Equivalente", vbInformation, "Mensaje"
        Text4.SetFocus: Exit Sub
    End If
        
       'Verificar si ya existe el codigo de Equivalencia con la UNidad respectiva
    If nOpc = 0 Then 'Ingreso
        
    Else            'MOdificacion
        If Trim(Text4) <> "" Then
            Label7 = fEqui(Text4, Text1)
            If Trim(Label7) <> "" Then
                MsgBox "Unidad de Medida Equivalente ya esta registrada", vbInformation, "Mensaje"
                Text4.SetFocus: Exit Sub
            End If
        End If
   End If
   
       Data2.Recordset.AddNew
       Data2.Recordset("EQUNIPRI") = Text1  'guardo el codigo
       Data2.Recordset("EQCANTEQUI") = Val(Text3)
       Data2.Recordset("EQUNIEQUI") = Text4  'item               guardo item
       Data2.Recordset.Update
       Data2.Refresh
       Frame3.Visible = False
       
 End If
 
 Frame1.Visible = True
 'If nOpc = 0 Then
 'If Trim(Text1) <> "" Then
  '      cUni = fUni(Text1)
   '     If Trim(cUni) = "" Then
    '        MsgBox "Unidad de Medida no registrada", vbInformation, "Mensaje"
     '       Text1.SetFocus: Exit Sub
      '  End If
  'Else
   '     MsgBox "Ingrese Unidad de Medida", vbInformation, "Mensaje"
    '    Text1.SetFocus: Exit Sub
 'End If
 'End If
 If Text2 = "" Then
    MsgBox "Ingrese Descripcion de Unidad de Medida", vbInformation, "Mensaje"
    Text2.SetFocus: Exit Sub
 End If
  If nOpc = 0 Then
    Data1.Recordset.AddNew
  Else
    Data1.Recordset.Edit
  End If
  
    Data1.Recordset("UM_ABREV") = Text1  'guardo el codigo
    Data1.Recordset("UM_NOMBRE") = Text2  'item               guardo item
   Data1.Recordset.Update
   Data1.Refresh
  Data1.Recordset.FindFirst "UM_ABREV ='" & Text1 & "'"
  
  If nOpc = 0 Then
     Limpiar
     Text1.SetFocus
  Else
     Frame1.Visible = False
     Frame5.Visible = True
     Frame2.Visible = False
     DBGrid1.Visible = True
     Frame4.Visible = True
     DBGrid1.SetFocus
  End If

Exit Sub
GrabErr:
    MsgBox Err.Description
    'If nTra = 1 Then cConexCom.RollbackTrans
End Sub

Private Sub Command21_Click()
If Frame3.Visible Then
   Frame3.Visible = False
   Frame1.Visible = True
   Frame2.Visible = True
   Text3 = ""
   Text4 = "": Label7 = ""
   If nOpc = 0 Then
   Text1.SetFocus
   Else
   Text2.SetFocus
   End If
   Exit Sub
End If
 If Frame1.Visible Then
   Frame1.Visible = False
   Frame2.Visible = False
   Limpiar
   Text1.Enabled = True
   Text2.Enabled = True
   DBGrid1.Visible = True
   Frame5.Visible = True
   Frame4.Visible = True
   Exit Sub
 End If
End Sub

Private Sub Command3_Click()
Dim cSql1 As String
Dim CSQL2 As String
Dim cCodigo1 As String
Dim pos As Variant
Dim cSel1 As Recordset
Dim cCodigo As String
Dim Mensaje$
On Error GoTo EliErr

If Data1.Recordset.RecordCount > 0 Then
    'Data1.RecordSource = "Select * from TABUNIMED order by UM_ABREV"
    'Data1.Refresh
    
    cSql1 = "Delete from TABUNIMED Where UM_ABREV= '" & Data1.Recordset("UM_ABREV") & "'"
    CSQL2 = "Delete from TABEQUI Where EQUNIPRI= '" & Data1.Recordset("UM_ABREV") & "'"
 
    
    If Data1.Recordset("UM_ESTADO") = "B" Then
      MsgBox "No se puede Eliminar dicha unidad", vbInformation, "Aviso"
      DBGrid1.SetFocus: Exit Sub
    End If
    
    Dim cSqlA As String, cSelA As ADODB.Recordset
    
    cSqlA = "Select * FROM TABEQUI WHERE EQUNIEQUI = '" & Trim(Data1.Recordset("UM_ABREV")) & "'"
    Set cSelA = New ADODB.Recordset
    cSelA.Open cSqlA, cConexCom, adOpenStatic
    If cSelA.RecordCount > 0 Then
       If MsgBox("La Unidad de Medida tiene registrada Medidas Equivalentes, desea Eliminarla de todas maneras", vbYesNo, "Eliminacion de Registro") = vbNo Then
          cSelA.Close: Exit Sub
       End If
    End If
    cSelA.Close
    

    If MsgBox("Seguro de Eliminar", vbQuestion + vbOKCancel, mensaje1) = vbOK Then
            nTra = 2
            cCodigo1 = Pos_Dato1(Data1.Recordset, "UM_ABREV")
            nTra = 1
            cConexCom.BeginTrans
            cConexCom.Execute CSQL2
            cConexCom.CommitTrans
            nTra = 0
            Data2.Refresh
            Data1.Refresh
            If cCodigo1 <> "" Then
            Data1.Recordset.FindFirst "UM_ABREV='" & cCodigo1 & "'"
            End If
            MsgBox "Datos Eliminados", vbInformation, "Mensaje"
            
    End If
    DBGrid1.Refresh
    
    If DBGrid1.Visible Then DBGrid1.SetFocus
Else
    MsgBox "No existe ningún registro para Eilminar", vbInformation, mensaje1
End If
Exit Sub
EliErr:
    MsgBox Err.Description
    If nTra = 1 Then cConexCom.RollbackTrans
End Sub

Private Sub Command4_Click()
  Dim src As String
  Dim ncar As String
  Dim criterio As String
   src = InputBox$("Ingrese el Nombre de la Unidad", "Búsqueda")
   ncar = Str$(Len(src))
   criterio = "Left(UM_NOMBRE," & ncar & ") = '" & src & "'"
   Data1.Recordset.FindFirst criterio
   If Data1.Recordset.NoMatch Then
      MsgBox "No se encontró el Registro !"
   End If
End Sub

Private Sub Command6_Click()
  Data1.RecordSource = "SELECT * FROM TabUniMed ORDER BY UM_NOMBRE"
  Data1.Refresh
End Sub

Private Sub Command7_Click()
   Unload Me
End Sub

Private Sub Command8_Click()
If Text1 <> "" And Frame1.Visible Then
    Data1.Recordset.AddNew
    Data1.Recordset(0) = Text1  'guardo el codigo
    Data1.Recordset(2) = "A"
   If Text2 <> "" Then
     Data1.Recordset(1) = Text2  'item               guardo item
     Data2.Recordset(0) = Text1  'guardo el codigo
     Data2.Recordset(2) = 1
     Data2.Recordset(1) = Text1  'item               guardo item
     Data2.Recordset.Update
     Data2.Refresh
   End If
  
   Data1.Recordset.Update
   Data1.Refresh
   Command1.Visible = True
   Command2.Visible = True
   Command3.Visible = True
   Command4.Visible = True
   command5.Visible = True
   Command6.Visible = True
   Command7.Visible = True
   Command8.Visible = False
   Command9.Visible = False
   Frame1.Visible = False
  
  End If
End Sub
Private Sub Form_Load()
Command20.Enabled = True
central Me         'Centra Formulario
'central Form3
Init_ControlDBGrid DBGrid1
Data1.DatabaseName = cRuta2
Data2.DatabaseName = cRuta2
Data1.RecordSource = "Select * from TABUNIMED order by UM_ABREV"
Data2.RecordSource = "Select * from TABEQUI order by EQUNIPRI,EQUNIEQUI"
End Sub

Private Sub Text1_GotFocus()
Enfoque Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
Dim cUni As String
If KeyAscii = 13 Then
   If Trim(Text1) <> "" Then
       If Existe(1, Text1, "TABUNIMED", "UM_ABREV", False) Then
         MsgBox "El código de Unidad ya existe", vbInformation, "Mensaje"
         Text3 = ""
         Text4 = ""
         Label7 = ""
         Text1 = "": Text1.SetFocus
      Else
          Text2.SetFocus
      End If
   Else
      MsgBox "Ingrese Código de Unidad de Medida", vbInformation, mensaje1
      Text1.SetFocus: Exit Sub
   End If
    
Else
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End If
End Sub

Private Sub Text2_GotFocus()
Enfoque Text2
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If Trim(Text2) <> "" Then
      Command20.SetFocus
   Else
      MsgBox "Ingrese Descripcion de Unidad de Medida", vbInformation, mensaje1
      Text2.SetFocus: Exit Sub
   End If
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
Dim i As Integer

If KeyAscii = 13 Then
   Command20.SetFocus
Else
  If (Chr$(KeyAscii) > "9" Or Chr$(KeyAscii) < "0") And Chr$(KeyAscii) <> "." And KeyAscii <> 8 Then
     KeyAscii = 0
  Else
     If Chr$(KeyAscii) = "." Then
        For i = 1 To Len(Text3)
            If Mid(Text3, i, 1) = "." Then KeyAscii = 0: Exit Sub
        Next
     End If
  End If
End If
End Sub

Private Sub Text4_DblClick()
 VGForm1 = 4
 Form1.Show 1
End Sub
Sub Limpiar()
Text1 = ""
Text2 = ""
Text3 = "0"
Text4 = ""
Label7 = ""
End Sub

Private Sub Text4_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then Text4_DblClick
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
Dim cUni As String
If KeyAscii = 13 Then
   If Trim(Text4) <> "" Then
       If Existe(1, Text4, "TABUNIMED", "UM_ABREV", False) = False Then
         MsgBox "Unidad de Medida no existe", vbInformation, "Mensaje"
         Text4 = "": Text4.SetFocus
      Else
          Text3.SetFocus
      End If
   Else
      MsgBox "Ingrese Código de Unidad de Medida Equivalente", vbInformation, mensaje1
      Text4.SetFocus: Exit Sub
   End If
        
       'Verificar si ya existe el codigo de Equivalencia con la UNidad respectiva
    If nOpc = 0 Then 'Ingreso
        
    Else            'MOdificacion
        If Trim(Text4) <> "" Then
            Label7 = fEqui(Text4, Text1)
            If Trim(Label7) <> "" Then
                MsgBox "Unidad de Medida Equivalente ya esta registrada", vbInformation, "Mensaje"
                Text4.SetFocus: Exit Sub
            End If
        End If
   End If
    
Else
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End If

End Sub
