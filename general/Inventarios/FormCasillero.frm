VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FormCasillero 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ubicación  en el Almacen"
   ClientHeight    =   4770
   ClientLeft      =   1575
   ClientTop       =   2670
   ClientWidth     =   7710
   Icon            =   "FormCasillero.frx":0000
   LinkTopic       =   "Form12"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4770
   ScaleWidth      =   7710
   Begin VB.CommandButton cmdsalirlogistica 
      Caption         =   "&Salir"
      Height          =   495
      Left            =   5348
      TabIndex        =   15
      Top             =   4170
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&Modificar"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1388
      TabIndex        =   14
      Top             =   4170
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Nuevo"
      Height          =   495
      Left            =   164
      TabIndex        =   3
      Top             =   4170
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Salir"
      Height          =   495
      Left            =   6572
      TabIndex        =   6
      Top             =   4170
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Eliminar"
      Height          =   495
      Left            =   3800
      TabIndex        =   5
      Top             =   4170
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2612
      TabIndex        =   4
      Top             =   4170
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   3975
      Left            =   72
      TabIndex        =   7
      Top             =   36
      Width           =   7488
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid 
         Height          =   3576
         Left            =   144
         TabIndex        =   16
         Top             =   288
         Width           =   4836
         _ExtentX        =   8520
         _ExtentY        =   6324
         _Version        =   393216
         Cols            =   6
         FixedCols       =   0
         RowHeightMin    =   240
         BackColorSel    =   -2147483643
         ForeColorSel    =   0
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         FocusRect       =   2
         ScrollBars      =   2
         Appearance      =   0
         FormatString    =   "^Cód Art|Descripción|Uni.|>Cantidad|>Saldo|>Cant.Recibida"
         RowSizingMode   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   6
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid list1 
         Height          =   1992
         Left            =   1296
         TabIndex        =   17
         Top             =   1656
         Width           =   3504
         _ExtentX        =   6191
         _ExtentY        =   3519
         _Version        =   393216
         Cols            =   6
         FixedCols       =   0
         RowHeightMin    =   240
         BackColorSel    =   -2147483643
         ForeColorSel    =   0
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         FocusRect       =   2
         ScrollBars      =   2
         Appearance      =   0
         FormatString    =   "^Cód Art|Descripción|Uni.|>Cantidad|>Saldo|>Cant.Recibida"
         RowSizingMode   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   6
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.ListBox List2 
         Height          =   1425
         Left            =   1320
         TabIndex        =   12
         Top             =   1680
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1320
         MaxLength       =   12
         TabIndex        =   2
         Text            =   "Text2"
         Top             =   1224
         Width           =   3492
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1320
         MaxLength       =   20
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label LBLNOM 
         Caption         =   "Label6"
         Height          =   264
         Left            =   2124
         TabIndex        =   13
         Top             =   1260
         Width           =   2640
      End
      Begin VB.Image Image1 
         Height          =   3516
         Left            =   4932
         Picture         =   "FormCasillero.frx":08CA
         Stretch         =   -1  'True
         Top             =   252
         Width           =   2400
      End
      Begin VB.Label Label5 
         Caption         =   "Ubicaciones"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Descripción"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label3"
         Height          =   288
         Left            =   1320
         TabIndex        =   1
         Top             =   720
         Width           =   3444
      End
      Begin VB.Label Label2 
         Caption         =   "Ubicación"
         Height          =   252
         Left            =   240
         TabIndex        =   9
         Top             =   1224
         Width           =   972
      End
      Begin VB.Label Label1 
         Caption         =   "Código"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   735
      End
   End
End
Attribute VB_Name = "FormCasillero"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim criterio As String
Dim RSQ As New ADODB.Recordset

Private Sub cmdsalirlogistica_Click()
Unload Me
End Sub

Private Sub Command1_Click()
 If Text2 <> "" And Text1 <> "" Then
   CASILLERO
 End If
End Sub

Private Sub Command2_Click()
  Dim rpta As Integer
  Text1 = Trim(Text1)
  If Text1 <> "" And List1.Row <> 0 Then
     Text2 = List1.TextMatrix(List1.Row, 0)
     If Text2 = "" Then Exit Sub
     rpta = MsgBox("Desea eliminar :" + Text2, vbExclamation + vbOKCancel, "Confirmacion")
     If rpta = vbOK Then
        criterio = "Delete from tabcasillero where TCODART = " & "'" + Text1.text + "'"
        criterio = criterio + " and TCODALM = " & "'" + VGAlma + "'"
        criterio = criterio + " and TCASILLERO = " & "'" + Trim(Text2) + "'"
        VGcnx.Execute criterio
        'LIST1.RemoveItem LIST1.Row
        agregarlista
        Text2 = ""
     End If
  End If
End Sub

Private Sub Command3_Click()
  If grid.Visible = False Then
     grid.Visible = True
     limpia
     LoadGrid
  Else
     Unload Me
  End If
End Sub

Private Sub Command4_Click()
 grid.Visible = False
 limpia
 Text1.SetFocus
End Sub

Private Sub command5_Click()
grid.Visible = False
limpia
Text1 = grid.TextMatrix(grid.Row, 0)
Label3 = grid.TextMatrix(grid.Row, 1)
agregarlista
End Sub

Private Sub Form_Activate()
  Text1.SetFocus
End Sub

'Sub listado(wcad)
'  Set DBGrid1.DataSource = Nothing
'  Set RS = Nothing
'
'  Set RS = Vgcnx.Execute(wcad)
'  Set DBGrid1.DataSource = RS
'  With DBGrid1
'      .Columns(0).Caption = "Codigo"
'      .Columns(0).Width = 1000
'      .Columns(1).Caption = "Descripcion"
'      .Columns(1).Width = 3800
'      .MarqueeStyle = dbgHighlightRow
'      .Refresh
'  End With
'
'End Sub

Private Sub Form_Load()
  central Me
  limpia
  'Data1.DatabaseName = cRuta2
  'Data2.DatabaseName = cRuta2
  'central FormCasillero
   Set RSQ = VGcnx.Execute("select * from tabcasillero")
   
   LoadGrid
End Sub

Private Sub CASILLERO()
 Dim rpta As Integer
 Dim rsbusca As New ADODB.Recordset
 
     criterio = "TCODART = " & "'" + Trim(Text1.text) + "'"
     criterio = criterio + " and TCODALM = " & "'" + VGAlma + "'"
     criterio = criterio + " and TCASILLERO = " & "'" + Trim(Text2) + "'"
     
      Set rsbusca = VGcnx.Execute("select * from tabcasillero where " & criterio)
      If rsbusca.RecordCount > 0 Then   ' Not
           MsgBox "Existe la ubicacion fisica", vbInformation, "AVISO"
           rsbusca.Close
           Exit Sub
      Else
            VGcnx.Execute "INSERT INTO tabcasillero " & _
                              "(TCODART,TCODALM,TCASILLERO)" & _
                              " VALUES(" & _
                              "'" & Trim(Text1) & "'," & _
                              "'" & VGAlma & "','" & Trim(Text2) & "')"
      End If
      rsbusca.Close
      Set rsbusca = Nothing
      
      List1.AddItem Text2 & Chr(9) & LBLNOM.Caption
      List1.SetFocus
      LoadGrid
End Sub

Public Sub limpia()
 Text1 = ""
 Text2 = ""
 Label3 = ""
 List1.Clear
 LBLNOM.Caption = ""
 If grid.Visible = False And Me.Visible = True Then
    Command1.Enabled = True
    Command2.Enabled = True
    Command4.Enabled = False
    command5.Enabled = False
 Else
    Command2.Enabled = False
    Command1.Enabled = False
    Command4.Enabled = True
    command5.Enabled = True
 End If
 
End Sub

Private Sub Text1_Change()
  Dim ncar As String
  Dim rsk As New ADODB.Recordset
  
  If Text1 <> "" Then
        ncar = Str$(Len(Text1.text))
        criterio = "Left(ACODIGO," & ncar & ") = '" & Text1.text & "'"
        Set rsk = VGcnx.Execute("select * from MAEART where " & criterio)
        If rsk.RecordCount > 0 Then
            Label3.Caption = cNull(rsk.Fields("ADESCRI"))
            'agregarlista
        Else
            'MsgBox "El codigo no existe", vbOKOnly, "No Encontrado"
            Label3 = ""
        End If
        rsk.Close
  Else
    Label3 = ""
  End If
End Sub

Private Sub Text1_DblClick()
   limpia
   VGForm1 = 10
   FormAyuArt1.Show 1
'   If Text1 <> "" And List1.Rows = 0 Then
       agregarlista
'   End If
End Sub

Private Sub Text1_GotFocus()
Dim rsl As New ADODB.Recordset

 If Text1 <> "" Then
         criterio = "ACODIGO = " & "'" + Text1.text + "'"
         Set rsl = VGcnx.Execute("select * from MAEART where " & criterio)
         If rsl.RecordCount > 0 Then
            Label3.Caption = cNull(rsl.Fields("ADESCRI"))
            'agregarlista
            Text2.SetFocus
         Else
            MsgBox "El Código no existe", vbOKOnly, "No Encontrado"
            Text1.SetFocus
         End If
         rsl.Close
     End If
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 46 Then
    Label3 = ""
    List1.Clear
 ElseIf KeyCode = 112 Then
      Text1_DblClick
 End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
Dim rsj As New ADODB.Recordset

     If KeyAscii = 13 And Text1 <> "" Then
         criterio = "ACODIGO = " & "'" + Text1.text + "'"
         Set rsj = VGcnx.Execute("select * from tabcasillero where " & criterio)
         If rsj.RecordCount > 0 Then
            Label3.Caption = rsj.Fields("ADESCRI")
            agregarlista
            Text2.SetFocus
         Else
            MsgBox "El codigo no existe", vbOKOnly, "No Encontrado"
            Text1.SetFocus
         End If
         rsj.Close
     End If
End Sub

Public Sub agregarlista()
  Dim rs As ADODB.Recordset
  Dim rsql As String
  rsql = "SELECT TABCASILLERO.TCASILLERO FROM TABCASILLERO WHERE TCODALM='" & VGAlma & "' and tcodart= '" & Text1 & "'"
  'rSql = "select tcasillero FROM tabcasillero where tcodalm= '" & VGAlma & "' and tcodart= '" & Text1 & "'" '
  Set rs = New ADODB.Recordset
  rs.Open rsql, VGcnx, adOpenStatic
  Set List1.DataSource = rs
  List1.FormatString = "<Descripción               "
  rs.Close
End Sub

Private Sub Text2_DblClick()
'Dim Adodc2 As ADODB.Recordset
'Set Adodc2 = New ADODB.Recordset
'Dim cBase As String
'cBase = cRuta2
''If UCase(Dir$(cBASE)) = UCase(cNomBd4) Then
'        Adodc2.Open "SELECT TABUBICA.COD_UBIC, TABUBICA.DESCRI FROM TABUBICA WHERE TABUBICA.COD_ALMA='" & VGAlma & "'", Vgcnx, adOpenStatic
'        frmReferencia.Conectar Adodc2, "SELECT TABUBICA.COD_UBIC, TABUBICA.DESCRI FROM TABUBICA WHERE TABUBICA.COD_ALMA='" & VGAlma & "'"
'        frmReferencia.Label1.Caption = "TABLA  DE  UBICACCIONES "
'        frmReferencia.show vbmodal
'        Adodc2.Close
'        If vGUtil(1) <> "" Then
'                Text2.text = (vGUtil(1))
'                LBLNOM.Caption = vGUtil(2)
'        End If
''End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 And Text2 <> "" Then
    Command1.SetFocus
  End If
End Sub
Public Sub LoadGrid()
Dim RSB As New ADODB.Recordset

Set RSB = New ADODB.Recordset
RSB.Open " SELECT  TABCASILLERO.TCODART , MAEART.ADESCRI" & _
        " FROM MAEART INNER JOIN TABCASILLERO ON MAEART.ACODIGO = TABCASILLERO.TCODART WHERE TCODALM='" & VGAlma & "' group by TABCASILLERO.TCODART,MAEART.ADESCRI ", VGcnx, adOpenForwardOnly, adLockReadOnly
Set grid.DataSource = RSB
grid.FormatString = "^Código             |<Descripción                                               "
RSB.Close
End Sub
