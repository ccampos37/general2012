VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FormConStk 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta de Stock de Articulos"
   ClientHeight    =   4965
   ClientLeft      =   660
   ClientTop       =   2070
   ClientWidth     =   9780
   Icon            =   "FormConStk.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   9780
   Begin VB.Frame Frame2 
      Caption         =   "Buscar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   4905
      Left            =   60
      TabIndex        =   23
      Top             =   30
      Width           =   9690
      Begin VB.CommandButton Command8 
         Caption         =   "&Retornar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   8310
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   180
         Width           =   1185
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Consulta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   7050
         MouseIcon       =   "FormConStk.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   180
         Width           =   1185
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   4170
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   345
         Width           =   2670
      End
      Begin VB.TextBox TxtBuscar 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1710
         TabIndex        =   0
         Top             =   345
         Width           =   1545
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "FormConStk.frx":0D0C
         Left            =   1710
         List            =   "FormConStk.frx":0D1C
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   735
         Width           =   1695
      End
      Begin MSFlexGridLib.MSFlexGrid FG1 
         Height          =   3135
         Left            =   120
         TabIndex        =   26
         Top             =   1350
         Width           =   9450
         _ExtentX        =   16669
         _ExtentY        =   5530
         _Version        =   393216
         Cols            =   3
         FixedCols       =   2
         BackColorSel    =   -2147483646
         GridLines       =   0
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Almacen :"
         Height          =   195
         Left            =   3390
         TabIndex        =   30
         Top             =   390
         Width           =   705
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "DobleClick para Serie o Lote"
         Height          =   195
         Left            =   7470
         TabIndex        =   28
         Top             =   4560
         Width           =   2025
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   120
         Picture         =   "FormConStk.frx":0D4A
         Top             =   360
         Width           =   480
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Cod. Articulo :"
         Height          =   195
         Left            =   660
         TabIndex        =   25
         Top             =   390
         Width           =   990
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Indice :"
         Height          =   195
         Left            =   675
         TabIndex        =   24
         Top             =   780
         Width           =   525
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos del Articulo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4140
      Left            =   360
      TabIndex        =   2
      Top             =   135
      Width           =   8640
      Begin VB.CommandButton Command1 
         Cancel          =   -1  'True
         Caption         =   "&Retornar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   7290
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   2910
         Width           =   1185
      End
      Begin VB.ListBox List1 
         Height          =   450
         Left            =   2280
         TabIndex        =   27
         Top             =   2520
         Width           =   1545
      End
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   2250
         TabIndex        =   22
         Top             =   3645
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   2250
         TabIndex        =   21
         Top             =   3285
         Width           =   1575
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   2280
         TabIndex        =   20
         Top             =   2160
         Width           =   1575
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   2280
         TabIndex        =   19
         Top             =   1800
         Width           =   1575
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   2280
         TabIndex        =   18
         Top             =   1440
         Width           =   1575
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   2280
         TabIndex        =   17
         Top             =   1080
         Width           =   615
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   2280
         TabIndex        =   16
         Top             =   720
         Width           =   4935
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   2280
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label9 
         Caption         =   "Cuenta Contable"
         ForeColor       =   &H80000007&
         Height          =   255
         Left            =   450
         TabIndex        =   11
         Top             =   3645
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label8 
         Caption         =   "Tipo de Articulo"
         ForeColor       =   &H80000007&
         Height          =   255
         Left            =   450
         TabIndex        =   10
         Top             =   3285
         Width           =   1575
      End
      Begin VB.Label Label7 
         Caption         =   "Ubicacion"
         ForeColor       =   &H80000007&
         Height          =   255
         Left            =   480
         TabIndex        =   9
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Grupo"
         ForeColor       =   &H80000007&
         Height          =   255
         Left            =   480
         TabIndex        =   8
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Linea"
         ForeColor       =   &H80000007&
         Height          =   255
         Left            =   480
         TabIndex        =   7
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Familia"
         ForeColor       =   &H80000007&
         Height          =   255
         Left            =   480
         TabIndex        =   6
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Unidad de Medidad"
         ForeColor       =   &H80000007&
         Height          =   255
         Left            =   480
         TabIndex        =   5
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Descripcion"
         ForeColor       =   &H80000007&
         Height          =   255
         Left            =   480
         TabIndex        =   4
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Codigo"
         ForeColor       =   &H80000007&
         Height          =   255
         Left            =   480
         TabIndex        =   3
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Label Label21 
      Caption         =   "Almacen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   2640
      TabIndex        =   14
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label20 
      Caption         =   "Central"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   3720
      TabIndex        =   13
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label19 
      Caption         =   "Label19"
      Height          =   15
      Left            =   3120
      TabIndex        =   12
      Top             =   360
      Width           =   15
   End
End
Attribute VB_Name = "FormConStk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ncar As String
Dim src As String
Dim almacen As String

Private Sub Carga_Almacen()
Dim rsql As String
Dim Rs As Recordset
Dim I As Integer
 
rsql = "select TAALMA,TADESCRI FROM TabAlm "
Set Rs = VGCNx.Execute(rsql)
While Not Rs.EOF
  Combo2.AddItem (Rs(0)) & "  " & (Rs(1))
  Rs.MoveNext
Wend

Rs.MoveFirst
For I = 0 To Rs.RecordCount - 1
  If Rs(0) = VGAlma Then
    Combo2.ListIndex = I
    Exit For
  Else
    Rs.MoveNext
  End If
Next
Rs.Close
End Sub


Private Sub Command1_Click()
 If Frame1.Visible Then
            Command3.Visible = True
            Frame1.Visible = False
            Frame2.Visible = True
            FG1.Visible = True
   Else
            Unload Me
   End If
'Dim criterio As String
'Dim criterio1 As String
'Dim rsa As New ADODB.Recordset
'
'src = InputBox$("Ingrese CODIGO", "Búsqueda")
'ncar = Str$(Len(src))
'
'criterio1 = "Left(STCODIGO," & ncar & ") = '" & src & "'"
''Data1.Recordset.FindFirst criterio1
'Set rsa = VGCNx.Execute("select * from stkart where " & criterio1)
'If rsa.RecordCount = 0 Then
'    MsgBox "No se encontró el Registro !", vbExclamation, mensaje1
'End If
'rsa.Close
'Set rsa = Nothing

End Sub

Private Sub Combo1_Click()
If Combo1.text = "Codigo" Then
     FG1.Col = Combo1.ListIndex
     FG1.Sort = 5
ElseIf Combo1.ListIndex = 3 Then
     FG1.Col = 4
     FG1.Sort = 5
Else
     FG1.Col = Combo1.ListIndex
     FG1.Sort = 5
End If
Label1.Caption = Combo1.text
End Sub

Private Sub Command2_Click()
   Dim criterio As String
   Dim criterio1 As String
   Dim src1 As String
 
   ncar = Str$(Len(src))
   criterio = "MID$(ACODIGO,1," + ncar + ") = '" & src & "'"
End Sub

Private Sub Combo2_Click()
almacen = Mid(Combo2, 1, 2)

'---------------------------
 
  Dim Rs As New ADODB.Recordset
  Dim rsql As String
  
  VGForm = 6
 
  rsql = "select  p.ACODIGO, p.ADESCRI,p.AFAMILIA,p.AUNIDAD, p.ACODIGO2,n.STSKDIS,n.STKPREPRO,n.STKPREULT,n.STKFECULT from MaeArt p, StkArt n where p.ACODIGO =  n.STCODIGO  and n.STALMA = '" & almacen & "' and n.stskdis >0 and p.afstock='1' ORDER BY ACODIGO,afamilia "
  
  Set Rs = New ADODB.Recordset
  Rs.Open rsql, VGCNx, adOpenDynamic, adLockReadOnly
  FG1.Row = 0
  Call llenagrilla(Rs)
  If Rs.RecordCount = 0 Then
       MsgBox "No hay articulos en este almacen", vbCritical
       Exit Sub
  End If
  Combo1.ListIndex = 0
TxtBuscar.text = ""
'---------------------------

End Sub

Private Sub Command3_Click()
  Dim criterio1 As String
  Dim src As String
  Dim RSB As New ADODB.Recordset
  
  Frame2.Visible = False
  Command3.Visible = False
  FG1.Visible = False
  Frame1.Visible = True
  src = FG1.TextMatrix(FG1.Row, 0)
  ncar = Str$(Len(src))
  criterio1 = "Left(ACODIGO," & ncar & ") = '" & src & "'"
  ' Data2.Recordset.FindFirst criterio1
  Set RSB = VGCNx.Execute("select  * from maeart where " & criterio1)
  If RSB.RecordCount = 0 Then
      MsgBox "No se encontró el Registro !", vbExclamation, "Aviso"
  Else
     Text1 = IIf(IsNull(RSB.Fields("ACODIGO")), "", RSB.Fields("ACODIGO"))
     Text2 = IIf(IsNull(RSB.Fields("ADESCRI")), "", RSB.Fields("ADESCRI"))
     Text3 = IIf(IsNull(RSB.Fields("AUNIDAD")), "", RSB.Fields("AUNIDAD"))
     Text4 = IIf(IsNull(RSB.Fields("AFAMILIA")), "", RSB.Fields("AFAMILIA"))
     If Not IsNull(RSB.Fields("AMODELO")) Then
            Text5 = RSB.Fields("AMODELO")
     End If
     If Not IsNull(RSB.Fields("AGRUPO")) Then
            Text6 = RSB.Fields("AGRUPO")
     End If
     agregacombo
     If Not IsNull(RSB.Fields("ATIPO")) Then
            Text8 = RSB.Fields("ATIPO")
     End If
     If Not IsNull(RSB.Fields("ACUENTA")) Then
            Text9 = RSB.Fields("ACUENTA")
     End If
  End If
  RSB.Close
  Set RSB = Nothing
  
End Sub

Private Sub Command8_Click()
   If Frame1.Visible Then
            Command3.Visible = True
            Frame1.Visible = False
            Frame2.Visible = True
            FG1.Visible = True
   Else
            Unload Me
   End If
End Sub

Private Sub FG1_DblClick()
VGForm1 = 30
VGcod = FG1.TextMatrix(FG1.Row, 0)
If Existe(1, VGcod, "MAEART", "ACODIGO", False, "S", "AFLOTE") Then
   FormAyuLote.Show 1
ElseIf Existe(1, VGcod, "MAEART", "ACODIGO", False, "S", "AFSERIE") Then
   FrmSeries.Show 1
End If
End Sub

Private Sub Form_Load()
  Dim Rs As New ADODB.Recordset
  Dim almacen As String
  Carga_Almacen
  
  Command3.Picture = MDIPrincipal.ImageList2.ListImages.item("Consultar").Picture
  Command8.Picture = MDIPrincipal.ImageList2.ListImages.item("Retornar").Picture
  Command1.Picture = MDIPrincipal.ImageList2.ListImages.item("Retornar").Picture

  Dim rsql As String
  VGForm = 6
 
  rsql = "select  p.ACODIGO, p.ADESCRI,p.AFAMILIA,p.AUNIDAD, n.STSKDIS,n.STKPREPRO,n.STKPREULT,n.STKFECULT from MaeArt p, StkArt n where p.ACODIGO =  n.STCODIGO and stalma='" & VGAlma & "' and n.stskdis >0 and p.afstock='1' ORDER BY ACODIGO,afamilia "
  Set Rs = New ADODB.Recordset
  Rs.Open rsql, VGCNx, adOpenDynamic, adLockReadOnly

  If Rs.RecordCount = 0 Then
       MsgBox "No hay articulos en este almacen", vbCritical
       Exit Sub
  End If
  central FormConStk
  FG1.Rows = 1

  Call llenagrilla(Rs)
  FG1.Visible = True
  Frame1.Visible = False
  Combo1.ListIndex = 0
  

End Sub


 Sub llenagrilla(Rs As ADODB.Recordset)
    FG1.Clear: FG1.Rows = 1
    FG1.FormatString = "   Codigo | Descripcion |  Cod. Fab.|  Unidad   |  Familia|   Stock| Costo Prom | Ult.Compra | Fec. Ult. Ing. "
    FG1.Row = 0
    FG1.ColWidth(0) = 800
    FG1.ColWidth(1) = 3500
    FG1.ColWidth(2) = 1
    FG1.ColWidth(3) = 890
    FG1.ColWidth(4) = 1
    FG1.ColWidth(5) = 1000
    FG1.ColWidth(6) = 1000
    FG1.ColWidth(7) = 1000
    FG1.ColWidth(8) = 1000
    FG1.AllowUserResizing = flexResizeColumns
    FG1.ColAlignment(1) = 1
    'FG1.ColAlignment(2) = 1
    If Rs.RecordCount > 0 Then
        Rs.MoveFirst
        FG1.Visible = False
        While Not Rs.EOF
              FG1.AddItem (Rs(0) & vbTab & Rs(1) & vbTab & Rs(2) & vbTab & Rs(3) & vbTab & Rs(4) & vbTab & Format(Rs(5), "####0.000") & vbTab & Format(IIf(Rs(5) = 0, 0, Rs(6)), "####0.000") & vbTab & Format(Rs(7), "####0.00"))  '& vbTab & Rs(8)
              Rs.MoveNext
        Wend
    End If
    FG1.Visible = True
 End Sub
 

Private Sub Txtbuscar_Change()
Dim I As Integer
Dim n As Integer
Dim Rs As New ADODB.Recordset
Dim rsql As String

   n = Combo1.ListIndex
   If n = 3 Then
     n = 4
   End If
   If TxtBuscar <> "" Then
       Select Case Combo1.ListIndex
        Case 0
            rsql = "select  p.ACODIGO, p.ADESCRI,p.AFAMILIA,p.AUNIDAD, p.ACODIGO2,n.STSKDIS,n.STKPREPRO,n.STKPREULT,n.STKFECULT from MaeArt p, StkArt n where p.ACODIGO =  n.STCODIGO and n.STALMA = '" & almacen & "' and left(p.ACODIGO," & Len(Trim(TxtBuscar)) & ")='" & TxtBuscar & "' and n.stskdis >0 and p.afstock='1' ORDER BY ACODIGO "
        Case 1
            rsql = "select  p.ACODIGO, p.ADESCRI,p.AFAMILIA,p.AUNIDAD, p.ACODIGO2,n.STSKDIS,n.STKPREPRO,n.STKPREULT,n.STKFECULT from MaeArt p, StkArt n where p.ACODIGO =  n.STCODIGO and n.STALMA = '" & almacen & "' and left(p.ADESCRI," & Len(Trim(TxtBuscar)) & ")='" & TxtBuscar & "' and n.stskdis >0 and p.afstock='1' ORDER BY ACODIGO "
        Case 2
            rsql = "select  p.ACODIGO, p.ADESCRI,p.AFAMILIA,p.AUNIDAD, p.ACODIGO2,n.STSKDIS,n.STKPREPRO,n.STKPREULT,n.STKFECULT from MaeArt p, StkArt n where p.ACODIGO =  n.STCODIGO and n.STALMA = '" & almacen & "' and left(p.AFAMILIA," & Len(Trim(TxtBuscar)) & ")='" & TxtBuscar & "' and n.stskdis >0 and p.afstock='1' ORDER BY ACODIGO "
        Case 3
            rsql = "select  p.ACODIGO, p.ADESCRI,p.AFAMILIA,p.AUNIDAD, p.ACODIGO2,n.STSKDIS,n.STKPREPRO,n.STKPREULT,n.STKFECULT from MaeArt p, StkArt n where p.ACODIGO =  n.STCODIGO and n.STALMA = '" & almacen & "' and left(p.ACODIGO2," & Len(Trim(TxtBuscar)) & ")='" & TxtBuscar & "' and n.stskdis >0 and p.afstock='1' ORDER BY ACODIGO "
       End Select
       
       Set Rs = VGCNx.Execute(rsql)
       Call llenagrilla(Rs)
       Rs.Close
   Else
       rsql = "select  p.ACODIGO, p.ADESCRI,p.AFAMILIA,p.AUNIDAD, p.ACODIGO2,n.STSKDIS,n.STKPREPRO,n.STKPREULT,n.STKFECULT from MaeArt p, StkArt n where p.ACODIGO =  n.STCODIGO and n.STALMA = '" & almacen & "' and n.stskdis >0  and p.afstock='1' ORDER BY ACODIGO "
       
       Set Rs = VGCNx.Execute(rsql)
       Call llenagrilla(Rs)
       Rs.Close
   End If
End Sub

Private Sub agregacombo()
Dim criterio As String
Dim adodc1 As ADODB.Recordset
Set adodc1 = New ADODB.Recordset
List1.Clear
criterio = "select * from tabcasillero  where  TCODALM = '" & VGAlma & "' AND TCODART = '" & FG1.TextMatrix(FG1.Row, 0) & " '  "
adodc1.Open criterio, VGCNx, adOpenStatic
If adodc1.RecordCount > 0 Then
  While Not adodc1.EOF
    List1.AddItem adodc1("TCASILLERO")
    adodc1.MoveNext
  Wend
End If
adodc1.Close

End Sub
