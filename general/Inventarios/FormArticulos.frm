VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form FormArticulos 
   Caption         =   "Articulos"
   ClientHeight    =   5760
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9285
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   5760
   ScaleWidth      =   9285
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Detalle"
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
      Height          =   4335
      Left            =   240
      TabIndex        =   30
      Top             =   240
      Width           =   8415
      Begin VB.ComboBox Combo4 
         Height          =   315
         ItemData        =   "FormArticulos.frx":0000
         Left            =   6000
         List            =   "FormArticulos.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   3720
         Width           =   735
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         ItemData        =   "FormArticulos.frx":0016
         Left            =   6000
         List            =   "FormArticulos.frx":0020
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   3240
         Width           =   735
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "FormArticulos.frx":002C
         Left            =   6000
         List            =   "FormArticulos.frx":002E
         TabIndex        =   43
         Text            =   "Combo2"
         Top             =   1320
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   2400
         TabIndex        =   2
         Top             =   840
         Width           =   4455
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "FormArticulos.frx":0030
         Left            =   6000
         List            =   "FormArticulos.frx":003D
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   2760
         Width           =   1695
      End
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   6000
         MaxLength       =   12
         TabIndex        =   9
         Top             =   2280
         Width           =   1695
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   2400
         TabIndex        =   8
         Top             =   3720
         Width           =   1695
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   2400
         TabIndex        =   7
         Top             =   3240
         Width           =   1695
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   2400
         TabIndex        =   6
         Top             =   2760
         Width           =   1695
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   2400
         MaxLength       =   6
         TabIndex        =   5
         Top             =   2280
         Width           =   495
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   2400
         TabIndex        =   4
         Top             =   1800
         Width           =   4455
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2400
         MaxLength       =   12
         TabIndex        =   3
         Top             =   1320
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2400
         TabIndex        =   1
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label lblunidad 
         Caption         =   "Label16"
         Height          =   255
         Left            =   3000
         TabIndex        =   55
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label Label12 
         Caption         =   "Stock x Lote"
         Height          =   255
         Left            =   4440
         TabIndex        =   45
         Top             =   3720
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Stock x Serie"
         Height          =   255
         Left            =   4440
         TabIndex        =   44
         Top             =   3240
         Width           =   1095
      End
      Begin VB.Label Label11 
         Caption         =   "Descripcion"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   40
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "Tipo de Articulo"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4440
         TabIndex        =   39
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "Ubicacion Almacen"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4440
         TabIndex        =   38
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "Grupo"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   37
         Top             =   3720
         Width           =   1455
      End
      Begin VB.Label lINE 
         Caption         =   "Linea"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   36
         Top             =   3240
         Width           =   1455
      End
      Begin VB.Label Label10 
         Caption         =   "Codigo Fab."
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   35
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label5 
         Caption         =   "Unidad Medida"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   34
         Top             =   2280
         Width           =   1935
      End
      Begin VB.Label Label4 
         Caption         =   "Familia"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   33
         Top             =   2760
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Descripcion Fab"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   32
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Codigo Interno"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   31
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Ficha Tecnica"
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
      Height          =   4335
      Left            =   240
      TabIndex        =   47
      Top             =   240
      Visible         =   0   'False
      Width           =   8415
      Begin VB.CommandButton Command21 
         Caption         =   "Cancelar"
         Height          =   615
         Left            =   7320
         TabIndex        =   50
         Top             =   2400
         Width           =   855
      End
      Begin VB.CommandButton Command20 
         Caption         =   "Grabar"
         Height          =   615
         Left            =   7320
         TabIndex        =   49
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox Text10 
         Height          =   3255
         Left            =   960
         MultiLine       =   -1  'True
         TabIndex        =   48
         Top             =   720
         Width           =   6015
      End
      Begin VB.Label Label15 
         Caption         =   "Label15"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4200
         TabIndex        =   54
         Top             =   360
         Width           =   3015
      End
      Begin VB.Label Label14 
         Caption         =   "CODIGO:"
         Height          =   255
         Left            =   960
         TabIndex        =   53
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label13 
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   52
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "DESCRIPCION:"
         Height          =   255
         Left            =   3000
         TabIndex        =   51
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.CommandButton Command7 
      Caption         =   "&Salir"
      Height          =   675
      Left            =   7200
      Picture         =   "FormArticulos.frx":0060
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   4800
      Width           =   775
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   "C:\Inventarios.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   8760
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "TabCasillero"
      Top             =   4320
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Inventarios.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   615
      Left            =   8040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5040
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&Imprimir"
      Height          =   675
      Left            =   6120
      Picture         =   "FormArticulos.frx":04A2
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   4800
      Width           =   775
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Buscar"
      Height          =   675
      Left            =   3960
      Picture         =   "FormArticulos.frx":08E4
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   4800
      Width           =   775
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Eliminar"
      Height          =   675
      Left            =   2880
      Picture         =   "FormArticulos.frx":0D26
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   4800
      Width           =   775
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Modificar"
      Height          =   675
      Left            =   1800
      Picture         =   "FormArticulos.frx":1168
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   4800
      Width           =   775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Adicionar"
      Height          =   675
      Left            =   720
      Picture         =   "FormArticulos.frx":15AA
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   4800
      Width           =   775
   End
   Begin VB.CommandButton Command6 
      Caption         =   "&Por Nombre"
      Height          =   675
      Left            =   7200
      Picture         =   "FormArticulos.frx":18B4
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   4800
      Visible         =   0   'False
      Width           =   775
   End
   Begin VB.CommandButton Command8 
      Caption         =   "&Grabar"
      Height          =   675
      Left            =   2880
      Picture         =   "FormArticulos.frx":1BBE
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   4800
      Visible         =   0   'False
      Width           =   775
   End
   Begin VB.CommandButton Command9 
      Caption         =   "&Cancelar"
      Height          =   675
      Left            =   6120
      Picture         =   "FormArticulos.frx":1EC8
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   4800
      Visible         =   0   'False
      Width           =   775
   End
   Begin VB.CommandButton Command10 
      Caption         =   "&Imprimir por Código"
      Height          =   675
      Left            =   1800
      Picture         =   "FormArticulos.frx":2012
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   4800
      Visible         =   0   'False
      Width           =   775
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Por Nombre&"
      Height          =   675
      Left            =   3960
      Picture         =   "FormArticulos.frx":231C
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   4800
      Visible         =   0   'False
      Width           =   775
   End
   Begin VB.CommandButton Command12 
      Caption         =   "&Cancelar Imprimir"
      Height          =   675
      Left            =   7200
      Picture         =   "FormArticulos.frx":2626
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   4815
      Visible         =   0   'False
      Width           =   775
   End
   Begin VB.CommandButton Command13 
      Caption         =   "&Ordenar por  Código"
      Height          =   675
      Left            =   1800
      Picture         =   "FormArticulos.frx":2770
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   4800
      Visible         =   0   'False
      Width           =   775
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Ordenar por &Nombre"
      Height          =   675
      Left            =   3960
      Picture         =   "FormArticulos.frx":28BA
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   4815
      Visible         =   0   'False
      Width           =   775
   End
   Begin VB.CommandButton Command15 
      Caption         =   "&Cancelar Ordenar"
      Height          =   675
      Left            =   7200
      Picture         =   "FormArticulos.frx":2BC4
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4815
      Visible         =   0   'False
      Width           =   775
   End
   Begin VB.CommandButton Command16 
      Caption         =   "&Buscar por Código"
      Height          =   675
      Left            =   1800
      Picture         =   "FormArticulos.frx":2D0E
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4800
      Visible         =   0   'False
      Width           =   775
   End
   Begin VB.CommandButton Command17 
      Caption         =   "Buscar por &Nombre"
      Height          =   675
      Left            =   3960
      Picture         =   "FormArticulos.frx":3150
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4815
      Visible         =   0   'False
      Width           =   775
   End
   Begin VB.CommandButton Command18 
      Caption         =   "&Cancelar Buscar"
      Height          =   675
      Left            =   7200
      Picture         =   "FormArticulos.frx":3592
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4815
      Visible         =   0   'False
      Width           =   775
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "FormArticulos.frx":36DC
      Height          =   3975
      Left            =   840
      OleObjectBlob   =   "FormArticulos.frx":36F0
      TabIndex        =   41
      Top             =   240
      Width           =   7455
   End
   Begin MSFlexGridLib.MSFlexGrid FG 
      Height          =   3855
      Left            =   0
      TabIndex        =   42
      Top             =   240
      Visible         =   0   'False
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   6800
      _Version        =   393216
   End
   Begin VB.CommandButton Command19 
      Caption         =   " &Ficha Tecnica"
      Height          =   675
      Left            =   5040
      Picture         =   "FormArticulos.frx":427B
      TabIndex        =   46
      Top             =   4800
      Width           =   775
   End
End
Attribute VB_Name = "FormArticulos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
   Dim Db As Database
   Dim agregar As Boolean
   Dim cod As String
   
Private Sub Command1_Click()

If Frame1.Visible Then
    limpia
Else
   limpia
   Command1.Visible = False
   Command2.Visible = False
   Command3.Visible = False
   Command4.Visible = False
   Command5.Visible = False
   Command6.Visible = False
   Command7.Visible = False
   Command19.Visible = False
   Command8.Visible = True
   Command9.Visible = True
   Frame1.Visible = True
   agregar = True
   Text1.SetFocus
End If

End Sub

Private Sub Command19_Click()
  Dim criterio As String
  cod = Data1.Recordset("ACODIGO")
  Text10 = ""
  Data1.RecordSource = "SELECT * FROM MAEART"
  Data1.Refresh
  criterio = "ACODIGO = " & "'" + cod + "'"
  Data1.Recordset.FindFirst criterio
  Label13 = Data1.Recordset("ACODIGO") 'cod
  Label15 = Data1.Recordset("ADESCRI")
  Text10 = IIf(Not IsNull(Data1.Recordset("ACOMENTA")), Data1.Recordset("ACOMENTA"), "")
  'Text10.SetFocus
  Deshabilitar (False)
  Frame2.Visible = True
  DBGrid1.Visible = False
  
End Sub

Private Sub Command2_Click()
  Dim precio As Double
  Dim cantidad As Double
  Dim usql As String
  Dim rsql As String
  Dim rs As Recordset
  Dim criterio As String
  Dim aux As String
  'Dim cod As String
  If Data1.Recordset.RecordCount = 0 Then Exit Sub
  limpia
  If Frame1.Visible Then
    
  Else
   Command1.Visible = False
   Command2.Visible = False
   Command3.Visible = False
   Command4.Visible = False
   Command5.Visible = False
   Command6.Visible = False
   Command7.Visible = False
   Command19.Visible = False
   Command8.Visible = True
   Command9.Visible = True
   Frame1.Visible = True
   agregar = False
   Text2.SetFocus
   cod = Data1.Recordset("ACODIGO")
  End If
  If Not Frame1.Visible Then
     'Text3.Text = "0"
     
     'usql = "Update MovArti set ADESCRI = " & Text3.Text & " where ACODIGO='" & Text1 & "'"
     'RS.Close
     'DB.Execute usql
     'Frame2.Visible = True
  End If
   Data1.RecordSource = "SELECT * FROM MAEART"
   Data1.Refresh
   criterio = "ACODIGO = " & "'" + cod + "'"
   Data1.Recordset.FindFirst criterio
         If Not Data1.Recordset.NoMatch Then
           'MsgBox "El Código de Articulo ya existe "
          End If
   Text1 = Data1.Recordset("ACODIGO")
   Text8 = Data1.Recordset("ADESCRI")
   If Not IsNull(Data1.Recordset("AFAMILIA")) Then
     Text5.text = Data1.Recordset("AFAMILIA")
   End If
   If Not IsNull(Data1.Recordset("AMODELO")) Then
      Text6.text = Data1.Recordset("AMODELO")
   End If
   If Not IsNull(Data1.Recordset("AGRUPO")) Then
     Text7.text = Data1.Recordset("AGRUPO")
   End If
   If Not IsNull(Data1.Recordset("AUNIDAD")) Then
     aux = Data1.Recordset("AUNIDAD")
     Text4 = unidad(aux)
     VGabrev = Text4
   End If
   If Not IsNull(Data1.Recordset("ACUENTA")) Then
       ' Text10.text = Data1.Recordset("ACUENTA")  se borro
   End If
   If Not IsNull(Data1.Recordset("ACODIGO2")) Then
      Text2.text = Data1.Recordset("ACODIGO2")
   End If
   
   If Not IsNull(Data1.Recordset("ADESCRI2")) Then
      
      Text3.text = Data1.Recordset("ADESCRI2")
   End If
  
      criterio = "TCODART = " & "'" + Text1.text + "'"
      criterio = criterio + " and TCODALM = " & "'" + VGAlma + "'"
      Data2.Recordset.FindFirst criterio
      If Not Data2.Recordset.NoMatch Then
        If Not IsNull(Data2.Recordset("TCASILLERO")) Then
              Text9.text = Data2.Recordset("TCASILLERO")
        End If
      End If
      rsql = "select e.TCASILLERO from TabCasillero e where  e.TCODALM = '" & VGAlma & "' AND e.TCODCOMP = '" & VGCOMP & "' and e.TCODART = '" & Text1.text & "'"
      Set Db = Workspaces(0).OpenDatabase(RUTA & NAMEBD)
      Set rs = Db.OpenRecordset(rsql, dbOpenSnapshot)
      If rs.EOF Then
        While Not rs.EOF
         Combo2.AddItem (rs(0))
         rs.MoveNext
        Wend
      End If
      If Not IsNull(Data1.Recordset("AFSERIE")) Then
        If Trim(Data1.Recordset("AFSERIE")) = "N" Then
            Combo3.ListIndex = 0
        Else
            Combo3.ListIndex = 1
        End If
    End If
    
    If Not IsNull(Data1.Recordset("AFLOTE")) Then
       If Trim(Data1.Recordset("AFLOTE")) = "N" Then
            Combo4.ListIndex = 0
        Else
           Combo4.ListIndex = 1
        End If
    End If
    
    
    
   ' If Not IsNull(Data1.Recordset("AUSER")) Then  'gUsu = Data1.Recordset("AUSER")
   'If Not IsNull(Data1.Recordset("AESTADO")) Then 'cEstado = Data1.Recordset("AESTADO")
    
    
  Frame1.Visible = True
  Text1.Enabled = False
  
End Sub

Private Sub Command20_Click()
 Dim rsql As String
 Set Db = Workspaces(0).OpenDatabase(RUTA & NAMEBD)
 
 rsql = "Update MaeART set ACOMENTA = '" & IIf(Text10 = "", " ", Text10) & "' "
 rsql = rsql & "Where ACODIGO = '" & cod & "' "
 Db.Execute rsql
 Frame2.Visible = False
 Deshabilitar (True)
 DBGrid1.Visible = True
End Sub

Private Sub Command21_Click()
  Frame2.Visible = False
  Deshabilitar (True)
  DBGrid1.Visible = True
  DBGrid1.Visible = True
End Sub

Private Sub Command3_Click()
  '  update   AESTADO= "A"
  Dim MENSAJE$
  Dim rsql As String
  Dim rs As Recordset
  Dim op  As Integer
  Dim codi As String
  If Data1.Recordset.RecordCount = 0 Then Exit Sub
  DBGrid1.Enabled = True
  DBGrid1.Visible = True
  Command1.Visible = True
  Command2.Visible = True
  Command3.Visible = True
  Command4.Visible = True
  Command5.Visible = True
  Command6.Visible = True
  Command7.Visible = True
  cod = Data1.Recordset("ACODIGO")
    op = Data1.Recordset.Bookmark
    rsql = "Select STCODIGO from StkArt where STALMA = '" & VGAlma & "' and  STCIA = '" & VGCOMP & "' and STCODIGO = '" & cod & "' "
    Set Db = Workspaces(0).OpenDatabase(RUTA & NAMEBD)
    Set rs = Db.OpenRecordset(rsql, dbOpenSnapshot)
    If rs.RecordCount > 0 Then
        MsgBox "El articulo tiene Movimiento de Almacen, no se puede Eliminar", vbInformation, "Mensaje"
        rs.Close
        Exit Sub
    End If
    rs.Close
   MENSAJE$ = "Está seguro de Borrar el Articulo ?" & Chr$(10)
   codi = Data1.Recordset("acodigo")
   MENSAJE$ = MENSAJE$ & codi
   If MsgBox(MENSAJE$, 33, "Borra el Articulo ?") <> 1 Then Exit Sub
   Db.Execute "delete from maeart where acodigo = '" & codi & " ' "
   'Data1.RecordSource = "SELECT * FROM MAEART "
'   Data1.Refresh
'   MENSAJE$ = "ACODIGO = " & "'" + cod + "'"
'    Data1.Recordset.FindFirst MENSAJE$
'     If Not Data1.Recordset.NoMatch Then
'          Data1.Recordset.Delete
'     End If
   
  ' Data1.RecordSource = "SELECT m.ACODIGO,m.ADESCRI ,m.AUNIDAD FROM MaeArt m,TabCasillero e  where   m.ACODIGO = e.TCODART  AND e.TCODCOMP = '" & VGComp & "' AND e.TCODALM = '" & VGAlma & "' GROUP BY m.ACODIGO, m.ADESCRI,m.AUNIDAD "
   Data1.Refresh
  ' Data1.Recordset.MoveNext
   If Not Data1.Recordset.EOF Then
        Data1.Recordset.MovePrevious
   Else
       Command7.SetFocus
   End If
End Sub

Private Sub Command4_Click()
  DBGrid1.Enabled = True
  DBGrid1.Visible = True
  Dim src As String
  Dim ncar As String
  Dim criterio As String
   src = InputBox$("Ingrese Codigo del Articulo", "Búsqueda")
   ncar = Str$(Len(src))
   criterio = "MID$(ACODIGO,1," + ncar + ") = " & Chr$(34) + src + Chr$(34)
   Data1.Recordset.FindFirst criterio
   If Data1.Recordset.NoMatch Then
      MsgBox "No se encontró el Registro !"
   End If
End Sub

Private Sub Command5_Click()
If Data1.Recordset.RecordCount = 0 Then Exit Sub
 MousePointer = vbHourglass
 
 If Frame1.Visible Then
   Frame1.Visible = False
   Command1.Visible = False
   Command2.Visible = False
   Command3.Visible = False
   Command4.Visible = False
   Command5.Visible = False
   Command6.Visible = False
   Command7.Visible = False
   
   Command10.Visible = True
   Command11.Visible = True
   Command12.Visible = True
 Else
   FormArtRep.Show 1
End If
 MousePointer = vbDefault
End Sub

Private Sub Command6_Click()
 
  Data1.RecordSource = "SELECT * FROM MAEART ORDER BY ADESCRI"
  Data1.Refresh
End Sub

Private Sub Command7_Click()
 If Frame1.Visible Then
   Command1.Visible = False
   Command2.Visible = False
   Command3.Visible = False
   Command4.Visible = False
   Command5.Visible = False
   Command6.Visible = False
   Command7.Visible = False
 Else
   Unload Me
 End If
End Sub

Private Sub Command8_Click()
 'Procedimiento para Grabar
Dim criterio As String
 lblunidad = ""
 If Text1 <> "" Then
    'grabar
   Data1.RecordSource = "SELECT * FROM MAEART"
   Data1.Refresh
   
    If agregar Then
     Data1.Recordset.AddNew
     Data1.Recordset(0) = Trim(Text1)  'guardo el codigo
     FG.AddItem (Text1 & vbTab & Text8)
    Else
     criterio = "ACODIGO = " & "'" + cod + "'"
     Data1.Recordset.FindFirst criterio
     If Not Data1.Recordset.NoMatch Then
           ' MsgBox "El Código de Articulo ya existe "
     End If
     Data1.Recordset.Edit
    End If
    
   If Text8 <> "" Then
     Data1.Recordset(2) = Trim(Text8)  'item               guardo item
   End If
   If Text2 <> "" Then
      Data1.Recordset(1) = Trim(Text2)  'guardo
   End If
   If Text3 <> "" Then
      Data1.Recordset(3) = Trim(Text3)  'guardo
   End If
   'Data1.Recordset("AHORA") = Time   EN LA BASE DE DB
   If Text4 <> "" Then
   'verificar si existe la unidad
      If agregar Or VGabrev = "" Then
        Data1.Recordset("AUNIDAD") = UCase(Trim(Text4))
      Else
      'VERIFICAR
        If Text4 <> VGabrev Then
         MsgBox " No se puede modificar la unidad", vbExclamation, "Aviso"
         Text4 = IIf(IsNull(Data1.Recordset("AUNIDAD")), " ", Data1.Recordset("AUNIDAD"))
         Exit Sub
        Else
         Data1.Recordset("AUNIDAD") = VGabrev
        End If
       End If
   End If
    VGabrev = ""
   If Text5 <> "" Then
      Data1.Recordset("AFAMILIA") = Text5
   End If
   If Text6 <> "" Then
      Data1.Recordset("AMODELO") = Text6
   End If
   If Text7 <> "" Then
      Data1.Recordset("AGRUPO") = Text7
   End If
    If Combo3.ListIndex = 0 Then
        Data1.Recordset("AFSERIE") = "N"
    Else
        Data1.Recordset("AFSERIE") = "S"
    End If
    If Combo4.ListIndex = 0 Then
        Data1.Recordset("AFLOTE") = "N"
    Else
        Data1.Recordset("AFLOTE") = "S"
    End If
    If (Combo3.ListIndex = 1) And (Combo4.ListIndex = 1) Then
        MsgBox "Error es Lote o Serie", vbCritical, "Error"
        Exit Sub
    End If
   'If Text10 <> "" Then
     ' Data1.Recordset("ACUENTA") = Text10
   'End If       se borro
   If Text9 <> "" Then
     criterio = "TCODART = " & "'" + Text1.text + "'"
     criterio = criterio + " and TCODALM = " & "'" + VGAlma + "'"
     criterio = criterio + " and TCODCOMP = " & "'" + VGCOMP + "'"
     criterio = criterio + " and TCASILLERO = " & "'" + Trim(Text9) + "'"
     Data2.Recordset.FindFirst criterio
      
      If Data2.Recordset.NoMatch Then
            Data2.Recordset.AddNew     'nuevo
            Data2.Recordset("TCODART") = Text1
            Data2.Recordset("TCODALM") = VGAlma
            Data2.Recordset("TCODCOMP") = VGCOMP
            Data2.Recordset("TCASILLERO") = Text9
            Data2.Recordset.Update
            Data2.Refresh
      End If
      
   End If
   
   If Combo1.ListIndex <> 0 Then
      Data1.Recordset("ATIPO") = "N"
   Else
      Data1.Recordset("ATIPO") = "I"
   End If
   Data1.Recordset.Update
   Data1.Refresh
   
   Command1.Visible = True
   Command2.Visible = True
   Command3.Visible = True
   Command4.Visible = True
   Command5.Visible = True
   Command6.Visible = True
   Command7.Visible = True
   Command19.Visible = True
   Command8.Visible = False
   Command9.Visible = False
   Frame1.Visible = False
   'FG.Visible = True
   VGcrea = False
  End If
  'muestra
End Sub

Private Sub Command9_Click()
If Frame1.Visible Then
   Command1.Visible = True
   Command2.Visible = True
   Command3.Visible = True
   Command4.Visible = True
   Command5.Visible = True
   Command6.Visible = True
   Command7.Visible = True
   Command8.Visible = False
   Command9.Visible = False
   Command19.Visible = True
   limpia
   
   Frame1.Visible = False
   Text1.Enabled = True
   
Else
   
  
End If

End Sub

Private Sub Form_Load()
  Dim Db As Database
  Dim rs As Recordset
  Dim rsql As String
  Text1.MaxLength = VGLongCodigo
  lblunidad = ""
  Data1.DatabaseName = RUTA & NAMEBD
  Data2.DatabaseName = RUTA & NAMEBD
  Init_ControlDBGrid DBGrid1
  central FormArticulos
  Combo2.Clear
  Combo1.ListIndex = 0
  Combo3.ListIndex = 0
  Combo4.ListIndex = 0
  Frame1.Visible = False

  Set Db = Workspaces(0).OpenDatabase(RUTA & NAMEBD)
  If VGcrea Then
     Command1.Visible = False
     Command2.Visible = False
     Command3.Visible = False
     Command4.Visible = False
     Command5.Visible = False
     Command6.Visible = False
     Command7.Visible = False
     Command8.Visible = True
     Command9.Visible = True
     Frame1.Visible = True
   End If
   'prueba
   Data1.RecordSource = "SELECT m.ACODIGO,m.ADESCRI ,m.AUNIDAD FROM MaeArt m GROUP BY m.ACODIGO, m.ADESCRI,m.AUNIDAD "   ',TabCasillero e  where   m.ACODIGO = e.TCODART  AND e.TCODCOMP = '" & VGComp & "' AND e.TCODALM = '" & VGAlma & "'
   Data1.Refresh
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Dim criterio As String
      
     If Len(Text1.text) = 8 Or Text1 <> "" Then
         criterio = "ACODIGO = " & "'" + Text1.text + "'"
         Data1.Recordset.FindFirst criterio
         If Not Data1.Recordset.NoMatch Then
            MsgBox "El Código de Articulo ya existe "
            Text1.SetFocus
         Else
            Text8.SetFocus
         End If
      End If
 End If
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
    SendKeys "{tab}"
    KeyAscii = 0
  End If

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
     Dim criterio As String
    If KeyAscii = 13 Then
        If Text2.text <> "" Then 'CODIGO DE FABRICANTE ES VARIABLE
         criterio = "ACODIGO2 = " & "'" + Text2.text + "'"
         'Data1.Recordset.FindFirst criterio
         'If Not Data1.Recordset.NoMatch Then
           ' MsgBox "El Código de Articulo ya existe "
           ' Text2.SetFocus
         'Else
            Text3.SetFocus
      Else
     ' End If  'si doy enter text4.set
         SendKeys "{tab}"
         KeyAscii = 0
      End If
 End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{tab}"
    KeyAscii = 0
  End If
End Sub

Private Sub Text4_DblClick()
  VGForm1 = 3
  Form1.Show 1
  Text4 = VGabrev
  Text5.SetFocus
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{tab}"
    KeyAscii = 0
  End If
End Sub

Private Sub Text5_DblClick()
   VGForm = 4
   'Form5.Show 1
End Sub
Private Sub limpia()
   Text1 = ""
   Text2 = ""
   Text3 = ""
   Text4 = ""
   Text5 = ""
   Text6 = ""
   Text7 = ""
   Text8 = ""
   Text9 = ""
   lblunidad = ""
   'Text10 = ""
   'Text11= ""
   'Combo1.Clear
   
End Sub

Public Sub prueba()
  
  Dim Rs1 As Recordset
  Dim Rsql1 As String
 

  Rsql1 = "select n.STCODIGO FROM  StkArt n where n.STALMA = '" & VGAlma & "'  and n.STCIA = '" & VGCOMP & "' "
  Set Rs1 = Db.OpenRecordset(Rsql1, dbOpenSnapshot)
  If Rs1.EOF Then
     MsgBox "No hay datos"
     Rs1.Close
     ' Close FormArticulos
     
  End If
End Sub

Public Sub grabar()
  Dim Db As Database
  Dim Rs1 As Recordset
  Dim insertar1 As String
  Dim Rsql1 As String
  
  Rsql1 = "select EP_CODART FROM  TabEmpArt  where EP_CODEMP= '" & VGCOMP & "'  AND EP_CODART = '" & Text1 & "' "
  Set Rs1 = Db.OpenRecordset(Rsql1, dbOpenSnapshot)
  If Not Rs1.EOF Then
     ' Rs1 (0)
     MsgBox "Codigo ya existe"
     
  Else
    insertar1 = "insert into TabEmpArt  values ('" & VGCOMP & "','" & Trim(Text1) & "' ) "
    Db.Execute insertar1
  End If
End Sub

Function unidad(aux As String) As String

 
 Dim rs As Recordset
 Dim rsql As String
  rsql = "select  UM_NOMBRE FROM TabUniMed where UM_ABREV= '" & aux & "'" '
  Set Db = Workspaces(0).OpenDatabase(RUTA & NAMEBD)
  Set rs = Db.OpenRecordset(rsql, dbOpenSnapshot)
  If Not rs.EOF Then
    unidad = rs(0)
  End If
  
End Function

Private Sub msf()
'   FG.FormatString = "Codigo Art.| Descripcion" '| Unidad|Cuenta |Grupo|Familia|Modelo| Ubicacion|Tipo "
'  FG.Row = 0
'  FG.ColWidth(0) = 1200
'  FG.ColWidth(1) = 4000
   'FG.ColWidth(2) = 1800
   '  Rsql = "SELECT m.ACODIGO , m.ADESCRI FROM MaeArt m ,TabCasillero e  where    e.TCODCOMP = '" & VGComp & "'  AND e.TCODALM = '" & VGAlma & "' AND m.ACODIGO = e.TCODART   GROUP BY m.ACODIGO ,m.ADESCRI"

'  Set Rs = Db.OpenRecordset(Rsql, dbOpenSnapshot)
'
'  FG.Rows = 1
'  Rs.MoveFirst
'  While Not Rs.EOF
'        FG.AddItem (Rs(0) & vbTab & Rs(1))    '& vbTab & RS(2))
'        Rs.MoveNext
'  Wend
'  Rs.Close
End Sub

Private Sub Deshabilitar(flag As Boolean)
   Command1.Enabled = flag
   Command2.Enabled = flag
   
   Command3.Enabled = flag
   Command4.Enabled = flag
   Command5.Enabled = flag
   Command6.Enabled = flag
   Command7.Enabled = flag
   Command8.Enabled = flag
   Command9.Enabled = flag
   Command11.Enabled = flag
   Command12.Enabled = flag
   Command13.Enabled = flag
   Command14.Enabled = flag
   Command15.Enabled = flag
   Command16.Enabled = flag
   Command17.Enabled = flag
   Command18.Enabled = flag
   Command19.Enabled = flag
  
   
   
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
    SendKeys "{tab}"
    KeyAscii = 0
  End If
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
   SendKeys "{tab}"
    KeyAscii = 0
  End If
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
  SendKeys "{tab}"
  KeyAscii = 0
End If
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{tab}"
    KeyAscii = 0
  End If
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
    SendKeys "{tab}"
    KeyAscii = 0
  End If
End Sub
