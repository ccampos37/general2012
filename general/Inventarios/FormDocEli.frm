VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form4 
   Caption         =   "FormDocumentos"
   ClientHeight    =   5925
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9375
   LinkTopic       =   "Form4"
   ScaleHeight     =   5925
   ScaleWidth      =   9375
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&Actualizar"
      Height          =   855
      Left            =   3960
      Picture         =   "FormDocEli.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   4680
      Width           =   1095
   End
   Begin VB.CommandButton Command7 
      Caption         =   "&Salir"
      Height          =   855
      Left            =   6480
      Picture         =   "FormDocEli.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Detalle del Documento"
      ForeColor       =   &H80000007&
      Height          =   4455
      Left            =   360
      TabIndex        =   5
      Top             =   120
      Width           =   8655
      Begin VB.TextBox Text1 
         BackColor       =   &H8000000B&
         Height          =   285
         Left            =   1800
         TabIndex        =   9
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H8000000B&
         Height          =   285
         Left            =   1800
         TabIndex        =   8
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1800
         TabIndex        =   7
         Top             =   1080
         Width           =   615
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   1800
         TabIndex        =   6
         Top             =   1440
         Width           =   615
      End
      Begin MSFlexGridLib.MSFlexGrid FG2 
         Height          =   2175
         Left            =   240
         TabIndex        =   10
         Top             =   2040
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   3836
         _Version        =   393216
         Rows            =   1
         Cols            =   7
         FormatString    =   "    CODIGO   |                DESCRIPCION            |      CANTIDAD      |  UNIDAD   |           PRECIO      |  UNI_ALM    "
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   285
         Left            =   6480
         TabIndex        =   24
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         Format          =   24838145
         CurrentDate     =   36437
         MaxDate         =   401768
         MinDate         =   36161
      End
      Begin VB.Label Label4 
         Caption         =   "Doc Referencial"
         Height          =   255
         Left            =   480
         TabIndex        =   20
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Proveedor"
         Height          =   255
         Left            =   480
         TabIndex        =   19
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Transaccion"
         Height          =   255
         Left            =   480
         TabIndex        =   18
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Documento"
         Height          =   255
         Left            =   480
         TabIndex        =   17
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label19 
         Caption         =   "Label1"
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
         Height          =   255
         Left            =   3360
         TabIndex        =   16
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label5 
         Caption         =   "Fecha"
         Height          =   255
         Left            =   5760
         TabIndex        =   15
         Top             =   360
         Width           =   735
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000009&
         BorderStyle     =   6  'Inside Solid
         X1              =   240
         X2              =   7920
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Label Label6 
         Caption         =   "Num"
         Height          =   255
         Left            =   2760
         TabIndex        =   14
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Lblndoc 
         Caption         =   "Label1"
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
         Height          =   255
         Left            =   2760
         TabIndex        =   13
         Top             =   1440
         Width           =   735
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
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   2760
         TabIndex        =   12
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label8 
         Caption         =   "Label8"
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
         Height          =   255
         Left            =   2760
         TabIndex        =   11
         Top             =   1080
         Width           =   1935
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Eliminar"
      Height          =   855
      Left            =   1680
      MouseIcon       =   "FormDocEli.frx":058C
      Picture         =   "FormDocEli.frx":09CE
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tipo de Documento"
      Height          =   4215
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   8775
      Begin VB.OptionButton Option1 
         Caption         =   "Nota de Ingreso"
         Height          =   375
         Left            =   2040
         TabIndex        =   3
         Top             =   600
         Width           =   3255
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Nota de Salida"
         Height          =   495
         Left            =   2040
         TabIndex        =   2
         Top             =   1440
         Width           =   2415
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Guias"
         Height          =   255
         Left            =   2040
         TabIndex        =   1
         Top             =   2280
         Width           =   1815
      End
   End
   Begin MSFlexGridLib.MSFlexGrid FG 
      Height          =   3975
      Left            =   480
      TabIndex        =   23
      Top             =   360
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   7011
      _Version        =   393216
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim db As Database
Dim estado As String
Dim Tipo As String * 2

Private Sub Command1_Click()
 Dim precio As Double
 
  Dim cantidad As Double
  Dim contador As Integer
  Dim Rsql1 As String
  Dim rsql As String
  Dim rs As Recordset
  Dim estado As String * 2
  If Frame2.Visible Then
    If Option3.Value Then
      Tipo = "GI"
     ElseIf Option2.Value Then
       Tipo = "NS"
     Else
       Tipo = "NI"
     End If
    estado = "E"
    rsql = "select  m.CATD, m.CANUMDOC, m.CACODMOV ,m.CAFECDOC, m.CACODPRO,m.CACODCLI, m.CARFTDOC ,m.CARFNDOC from MovAlmCab m where  m.CAALMA ='" & VGAlma & "' and m.CATD='" & Tipo & "' and m.CASITGUI<>'" & estado & "'  ORDER BY m.CANUMDOC" '
    Set db = Workspaces(0).OpenDatabase(RUTA & NAMEBD)
    
    Set rs = db.OpenRecordset(rsql, dbOpenSnapshot)
    FG.Rows = 1
    rs.MoveFirst
    If rs.EOF Then
       MsgBox "no hay datos", vbCritical, "Aviso"
       Exit Sub
    End If
    While Not rs.EOF
        FG.AddItem (rs(0) & vbTab & rs(1) & vbTab & rs(2) & vbTab & rs(3) & vbTab & rs(4) & vbTab & rs(5) & vbTab & rs(6) & vbTab & rs(7))
        rs.MoveNext
    Wend
    rs.Close
    Frame2.Visible = False
    Command1.Caption = "Validar"
    Command2.Visible = True
    Exit Sub
  End If
  If Frame1.Visible Then
  
     'Text3.Text = "0"
    
     Frame1.Visible = False
     
  Else
     Command1.Visible = False
     Frame1.Visible = True
      Set db = Workspaces(0).OpenDatabase(RUTA & NAMEBD)
     Command1.Caption = "&Aceptar"
     Command2.Visible = False
     '"Tipo Doc.|Numero de Doc| Tr| Fecha | Proveedor|Cliente|Td REF|Num.Doc Ref."
     DTPicker1 = FG.TextMatrix(FG.Row, 3)  'fecha
     Text2 = FG.TextMatrix(FG.Row, 2)  'tras
     Label7 = transa
     Text1 = FG.TextMatrix(FG.Row, 0) ' tipo de doc
     Label19 = FG.TextMatrix(FG.Row, 1) ' cod de doc
     Text3 = FG.TextMatrix(FG.Row, 4)  'proveedor
     Label8 = prove
     Text4 = FG.TextMatrix(FG.Row, 6)  'doc ref
     Lblndoc = FG.TextMatrix(FG.Row, 7)  'proveedor
     'Rsql = "select  from TabTransa  where  n.DEALMA ='"  '
     'Set Rs = Db.OpenRecordset(Rsql, dbOpenDynaset)
     'label7= Rs(0)
     ' MSFlexGrid1.TextMatrix(0, 1) = " DESCRIPCION"
     'MSFlexGrid1.TextMatrix(0, 2) = " CANTIDAD ING"
     'MSFlexGrid1.TextMatrix(0, 3) = " UNIDAD ING"
     'MSFlexGrid1.TextMatrix(0, 4) = " PRECIO UNIT"
     'MSFlexGrid1.TextMatrix(0, 5) = " CANT INF"
     'MSFlexGrid1.TextMatrix(0, 6) = " PRECIO INF"
     Rsql1 = "select n.DECODIGO, m.ADESCRI, m.AUNIDAD, n.DECANTID, n.DEPRECIO  from MovAlmDet n, MaeArt m where  n.DEALMA ='" & VGAlma & "' AND n.DETD = '" & Text1 & "' AND n.DENUMDOC ='" & Label19 & "' AND m.ACODIGO = n.DECODIGO ORDER BY n.DEITEM "  '
     Set rs = db.OpenRecordset(Rsql1, dbOpenDynaset)
     If rs.EOF Then
       Exit Sub
     End If
     rs.MoveFirst
     FG2.Rows = 1
    
     While Not rs.EOF
       FG2.AddItem (rs(0) & vbTab & rs(1) & vbTab & rs(2) & vbTab & rs(3) & vbTab & rs(4))
       rs.MoveNext
     Wend
     'cantidad = FG2.Rows
     rs.Close
  End If
End Sub

Private Sub Command2_Click()
 
 Dim usql As String
   estado = "E"  'eliminado
   'eliminar
   usql = "Update MovAlmCab set CASITGUI= " & estado & " where CANUMDOC='" & Trim(FG.TextMatrix(FG.Row, 1)) & "' and CATD ='" & Tipo & "'  and CAALMA='" & VGAlma & "'"
   db.Execute usql
   
End Sub

Private Sub Command7_Click()
  If Frame1.Visible Then
     limpia
     Frame1.Visible = False
     Command1.Visible = True
     Command2.Visible = True
  Else
     db.Close
     Unload Me
   End If
End Sub

Private Sub Form_Load()
 Dim db As Database
 Dim rs As Recordset
 Dim rsql As String
 'Dim Rsql As String
  limpia
  Label7 = ""
  Label8 = ""
  Command1.Caption = "Aceptar"
  Command2.Visible = False
  central FormConValArt
  FG.FormatString = "Tipo Doc.|Numero de Doc| Tr| Fecha | Proveedor|Cliente|Td REF|Num.Doc Ref."
  FG.Row = 0
  FG.ColWidth(0) = 800
  FG.ColWidth(1) = 1500
  FG.ColWidth(2) = 800
  FG.ColWidth(3) = 1000
  FG.ColWidth(4) = 1000
  FG.ColWidth(5) = 1000
  FG.ColWidth(6) = 800
  FG.ColWidth(7) = 1500
  
 
  Frame1.Visible = False
  'Rsql = "select  p.ACODIGO, p.ADESCRI, m.CACODMOV ,n.DENUMDOC, m.CANOMPRO, m.CARFTDOC from MaeArt p,MovAlmCab m, MovAlmDet n where p.ACODIGO = n.DECODIGO and n.DEPRE <> 0 and m.CANUMDOC= n.DENUMDOC ORDER BY m.CANUMDOC"
  'Set Db = Workspaces(0).OpenDatabase(RUTA & NAMEBD)
  'Set Rs = Db.OpenRecordset(Rsql, dbOpenSnapshot)
End Sub

Public Sub limpia()
  
  Text1 = ""
  Text2 = ""
  Text3 = ""
  'Label14 = ""
 ' Label16 = ""
  'Label18 = ""
  Label19 = ""
  FG2.Clear
  FG2.Rows = 1
  FG2.TextMatrix(0, 0) = " CODIGO"
  FG2.TextMatrix(0, 1) = " DESCRIPCION"
  FG2.TextMatrix(0, 2) = " CANTIDAD ING"
  FG2.TextMatrix(0, 3) = " UNIDAD ING"
  FG2.TextMatrix(0, 4) = " PRECIO UNIT"
  FG2.TextMatrix(0, 5) = " CANT INF"
  FG2.TextMatrix(0, 6) = " PRECIO INF"
  
  
  
End Sub

Function transa() As String

 Dim db As Database
 Dim rs As Recordset
 Dim rsql As String
  rsql = "select  TT_DESCRI FROM TabTransa where TT_CODMOV= '" & FG.TextMatrix(FG.Row, 2) & "'" '
  Set db = Workspaces(0).OpenDatabase(RUTA & NAMEBD)
  Set rs = db.OpenRecordset(rsql, dbOpenSnapshot)
  If Not rs.EOF Then
    transa = rs(0)
  End If
  
End Function
Function prove() As String

 Dim db As Database
 Dim rs As Recordset
 Dim rsql As String
  rsql = "select PRVCNOMBRE FROM maeprov where PRVCCODIGO= '" & FG.TextMatrix(FG.Row, 4) & "'" '
  Set db = Workspaces(0).OpenDatabase(RUTA & NAMEBD)
  Set rs = db.OpenRecordset(rsql, dbOpenSnapshot)
  If Not rs.EOF Then
    prove = rs(0)
  End If
End Function


'Private Sub eliminar()
' Dim Db As Database
' Dim Rs As Recordset
' Dim Rsql As String
'  Rsql = "select  m.CATD, m.CANUMDOC, m.CACODMOV ,m.CAFECDOC, m.CACODPRO,m.CACODCLI, m.CARFTDOC ,m.CARFNDOC from MovAlmCab m where   m.CAALMA ='" & VGAlma & "' and m.CATD='" & tipo & "' and m.CAESTIMP <>'" & ESTADO & "'  ORDER BY m.CANUMDOC" '
'  Set Db = Workspaces(0).OpenDatabase(RUTA & NAMEBD)
'  Set Rs = Db.OpenRecordset(Rsql, dbOpenSnapshot)
'  Rs.MoveFirst
'
'
'
'End Sub


