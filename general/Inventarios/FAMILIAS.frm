VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form Form5 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Familias"
   ClientHeight    =   5895
   ClientLeft      =   240
   ClientTop       =   1005
   ClientWidth     =   9045
   ControlBox      =   0   'False
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleMode       =   0  'User
   ScaleWidth      =   9135
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1095
      Left            =   240
      TabIndex        =   27
      Top             =   4680
      Visible         =   0   'False
      Width           =   8535
      Begin VB.CommandButton Command21 
         Caption         =   "&Salir"
         Height          =   735
         Left            =   4920
         Picture         =   "FAMILIAS.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   240
         Width           =   835
      End
      Begin VB.CommandButton Command20 
         Caption         =   "&Aceptar"
         Height          =   735
         Left            =   2520
         Picture         =   "FAMILIAS.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   240
         Width           =   835
      End
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   7080
      TabIndex        =   26
      Top             =   1560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   7080
      TabIndex        =   25
      Top             =   1200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command19 
      Caption         =   "&Líneas"
      Height          =   735
      Left            =   7200
      Picture         =   "FAMILIAS.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Command18 
      Caption         =   "&Cancelar Búsqueda"
      Height          =   735
      Left            =   6480
      Picture         =   "FAMILIAS.frx":3026
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   4800
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Command17 
      Caption         =   "Buscar  &Familia"
      Height          =   735
      Left            =   4080
      Picture         =   "FAMILIAS.frx":3170
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   4800
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Command16 
      Caption         =   "&Buscar Código"
      Height          =   735
      Left            =   1680
      Picture         =   "FAMILIAS.frx":35B2
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   4800
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Command15 
      Caption         =   "Cancelar Orden"
      Height          =   735
      Left            =   6480
      Picture         =   "FAMILIAS.frx":39F4
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   4800
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Ordenar por &Familia"
      Height          =   735
      Left            =   4080
      Picture         =   "FAMILIAS.frx":3B3E
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   4800
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Command13 
      Caption         =   "&Ordenar por Código"
      Height          =   735
      Left            =   1680
      Picture         =   "FAMILIAS.frx":3E48
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   4800
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Command12 
      Caption         =   "&Cancelar Imprimir"
      Height          =   735
      Left            =   6480
      Picture         =   "FAMILIAS.frx":3F92
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   4800
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Imprimir por &Familia"
      Height          =   735
      Left            =   4080
      Picture         =   "FAMILIAS.frx":40DC
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   4800
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Command10 
      Caption         =   "&Imprimir por Código"
      Height          =   735
      Left            =   1680
      Picture         =   "FAMILIAS.frx":43E6
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4800
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Command9 
      Caption         =   "&Cancelar"
      Height          =   735
      Left            =   5280
      Picture         =   "FAMILIAS.frx":46F0
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4800
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2760
      TabIndex        =   13
      Top             =   3240
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2760
      TabIndex        =   11
      Top             =   2040
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton Command8 
      Caption         =   "&Grabar"
      Height          =   735
      Left            =   2880
      Picture         =   "FAMILIAS.frx":4B32
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4800
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Inventarios.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   7200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "FAMILIA"
      Top             =   120
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton Command7 
      Caption         =   "&Salir"
      Height          =   735
      Left            =   7680
      Picture         =   "FAMILIAS.frx":4F74
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4800
      Width           =   855
   End
   Begin VB.CommandButton Command6 
      Caption         =   "&Ordenar"
      Height          =   735
      Left            =   6480
      Picture         =   "FAMILIAS.frx":50BE
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4800
      Width           =   855
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&Imprimir"
      Height          =   735
      Left            =   5280
      Picture         =   "FAMILIAS.frx":53C8
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4800
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Buscar"
      Height          =   735
      Left            =   4080
      Picture         =   "FAMILIAS.frx":580A
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4800
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Eliminar"
      Height          =   735
      Left            =   2880
      Picture         =   "FAMILIAS.frx":5C4C
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4800
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Modificar"
      Height          =   735
      Left            =   1680
      Picture         =   "FAMILIAS.frx":608E
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4800
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Adicionar"
      Height          =   735
      Left            =   600
      Picture         =   "FAMILIAS.frx":64D0
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4800
      Width           =   855
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "FAMILIAS.frx":67DA
      Height          =   3450
      Left            =   120
      OleObjectBlob   =   "FAMILIAS.frx":67EE
      TabIndex        =   1
      Top             =   1080
      Width           =   8295
   End
   Begin VB.Label Label3 
      Caption         =   "Nombre:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   12
      Top             =   3240
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Código:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   10
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "FAMILIAS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   3840
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim resp As String
Private Sub Command1_Click()
   resp = "S"

 Command19.Visible = False
 
   DBGrid1.Visible = False
   Command1.Visible = False
   Command2.Visible = False
   Command3.Visible = False
   Command4.Visible = False
   Command5.Visible = False
   Command6.Visible = False
   Command7.Visible = False
   
   Command8.Visible = True
   Command9.Visible = True
   
   Label1.Visible = True
   Label2.Visible = True
  
   
  
   Text1.Visible = True
   Text2.Visible = True
  
Text1.SetFocus
End Sub
Private Sub Command10_Click()
   Data1.RecordSource = "SELECT * FROM FAMILIA ORDER BY FAM_CODIGO"
   Data1.Refresh
End Sub
Private Sub Command11_Click()
   Data1.RecordSource = "SELECT * FROM FAMILIA ORDER BY FAM_CODIGO"
   Data1.Refresh
End Sub

Private Sub Command12_Click()
   DBGrid1.Visible = True
   Command1.Visible = True
   Command2.Visible = True
   Command3.Visible = True
   Command4.Visible = True
   Command5.Visible = True
   Command6.Visible = True
   Command7.Visible = True
   Command10.Visible = False
   Command11.Visible = False
   Command12.Visible = False
End Sub

Private Sub Command16_Click()
 
DBGrid1.Enabled = True
DBGrid1.Visible = True

src = InputBox$("Ingrese Código de Familia", "Búsqueda")
ncar = Str$(Len(src))
criterio = "MID$(FAM_CODIGO,1," + ncar + ") = " & Chr$(34) + src + Chr$(34)
Data1.Recordset.FindFirst criterio
If Data1.Recordset.NoMatch Then
   MsgBox "No se encontró el Registro !"
End If



End Sub
Private Sub Command17_Click()
DBGrid1.Enabled = True
DBGrid1.Visible = True
src = InputBox$("Ingrese Familia ", "Búsqueda")
  ncar = Str$(Len(src))
  criterio = "MID$(Fam_Nombre,1," + ncar + ") = " & Chr$(34) + src + Chr$(34)
  Data1.Recordset.FindFirst criterio
   If Data1.Recordset.NoMatch Then
   MsgBox "No se encontró el Registro !"
   End If
End Sub
Private Sub Command18_Click()
DBGrid1.Visible = True
   Command1.Visible = True
   Command2.Visible = True
   Command3.Visible = True
   Command4.Visible = True
   Command5.Visible = True
   Command6.Visible = True
   Command7.Visible = True
   Command16.Visible = False
   Command17.Visible = False
   Command18.Visible = False
   Command19.Visible = True
  End Sub

Private Sub Command19_Click()
Text3.text = Data1.Recordset(0)
Text4.text = Data1.Recordset(1)
FormArticulos.Text5 = Data1.Recordset(0)
Frame1.Visible = True
'Command20.Visible = True
Form6.Show 1
Unload Me
End Sub

Private Sub Command2_Click()
   resp = "N"
   DBGrid1.Visible = False
  
   Command1.Visible = False
   Command2.Visible = False
   Command3.Visible = False
   Command4.Visible = False
   Command5.Visible = False
   Command6.Visible = False
   Command7.Visible = False
   
   Command8.Visible = True
   Command9.Visible = True
   
   Label1.Visible = True
   Label2.Visible = True

   Text1.text = Data1.Recordset(0)
  
    If Not IsNull(Data1.Recordset(1)) Then
      Text2.text = Data1.Recordset(1)
    Else
      Text2.text = ""
    End If
    
   Text1.Visible = True
   Text1.Enabled = False
   Text2.Visible = True
   Text2.SetFocus
 
    Command19.Visible = False

End Sub

Private Sub Command20_Click()
   If Not IsNull(Data1.Recordset(0)) Then
     FormArticulos.Text5 = Data1.Recordset(0)
   End If
   Frame1.Visible = True
   Unload Me
End Sub

Private Sub Command21_Click()
   Frame1.Visible = True
   Unload Me
End Sub

Private Sub Command3_Click()
   
   Dim MENSAJE$
   
  DBGrid1.Enabled = True
  DBGrid1.Visible = True
  
Label1.Visible = True

 Command1.Visible = True
 Command2.Visible = True
 Command3.Visible = True
 Command4.Visible = True
 Command5.Visible = True
 Command6.Visible = True
 Command7.Visible = True
       
   MENSAJE$ = "Está seguro de Borrar la Familia?" & Chr$(10)
   MENSAJE$ = MENSAJE$ & Data1.Recordset(1)
   If MsgBox(MENSAJE$, 33, "Borra la Familia ?") <> 1 Then Exit Sub
   Data1.Recordset.Delete
   Data1.Recordset.MoveNext
   If Data1.Recordset.EOF Then Data1.Recordset.MovePrevious
 
DBGrid1.Visible = True

End Sub

Private Sub Command4_Click()
DBGrid1.Visible = True
DBGrid1.Enabled = True

Command1.Visible = False
Command2.Visible = False
Command3.Visible = False
Command4.Visible = False
Command5.Visible = False
Command6.Visible = False
Command7.Visible = False

Command16.Visible = True
Command17.Visible = True
Command18.Visible = True

Command19.Visible = False

End Sub

Private Sub Command5_Click()
   
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
   Command19.Visible = False
  
End Sub

Private Sub Command6_Click()
  
   Command1.Visible = False
   Command2.Visible = False
   Command3.Visible = False
   Command4.Visible = False
   Command5.Visible = False
   Command6.Visible = False
   Command7.Visible = False
     
   Command13.Visible = True
   Command14.Visible = True
   Command15.Visible = True
   
   Command19.Visible = False
   
End Sub

Private Sub Command13_Click()
   Data1.RecordSource = "Select * from FAMILIA order by FAM_CODIGO"
   Data1.Refresh
End Sub
Private Sub Command7_Click()
   Unload Form5
 End Sub

Private Sub Command8_Click()
  If Text1 = "" Then
     MsgBox "Ingrese codigo", vbExclamation, "Aviso"
     Text1.SetFocus
     Exit Sub
  End If
   Label1.Visible = True
   
   Command1.Visible = True
   Command2.Visible = True
   Command3.Visible = True
   Command4.Visible = True
   Command5.Visible = True
   Command6.Visible = True
   Command7.Visible = True
   
   Command8.Visible = False
   Command9.Visible = False
       
      
If Not IsNull(Text1.text) And Text1.text <> "" Then
      If resp = "S" Then
         Data1.Recordset.AddNew
         Data1.Recordset(0) = UCase$(Text1.text)
      Else
         Data1.Recordset.Edit
      End If
      
    If Not IsNull(Text2.text) Then
         Data1.Recordset(1) = Mid$(UCase$(Text2.text), 1, 45)
      Else
         Data1.Recordset(1) = " "
      End If
      Data1.Recordset.Update
   Data1.Refresh

    
End If
            
   
   Label1.Visible = False
   Label2.Visible = False
   
   Text1.Enabled = True
   
   Text1.Visible = False
   Text2.Visible = False
    
   Text1 = ""
   Text2 = ""
  
   
 DBGrid1.Visible = True
 DBGrid1.Enabled = True
 
 Command19.Visible = True
 
 
End Sub

Private Sub Command15_Click()
   Command1.Visible = True
   Command2.Visible = True
   Command3.Visible = True
   Command4.Visible = True
   Command5.Visible = True
   Command6.Visible = True
   Command7.Visible = True
   
   Command13.Visible = False
   Command14.Visible = False
   Command15.Visible = False
   
   Command19.Visible = True

  
End Sub

Private Sub Command14_Click()
   Data1.RecordSource = "Select * from FAMILIA order by FAM_NOMBRE"
   Data1.Refresh
End Sub

Private Sub Command9_Click()
   DBGrid1.Visible = True
   
   Command19.Visible = True
   
   Command1.Visible = True
   Command2.Visible = True
   Command3.Visible = True
   Command4.Visible = True
   Command5.Visible = True
   Command6.Visible = True
   Command7.Visible = True
   
   Command8.Visible = False
   Command9.Visible = False
 
   Label1.Visible = True
   
   Text1.Enabled = True
   Text1.Visible = False
   Text2.Visible = False
  
   Text1 = ""
   Text2 = ""
  
End Sub

Private Sub Form_Load()
Data1.DatabaseName = RUTA & NAMEBD
If VGForm = 4 Then
    Frame1.Visible = True
End If
Init_ControlDataGrid DBGrid1
Label1.Visible = True
Label2.Visible = True

Text3.Enabled = True
Text4.Enabled = True

Command19.Visible = True
central Form5
End Sub

Private Sub Text1_Change()
    If resp = "S" Then
      If Len(Text1.text) = 4 Then
         criterio = "FAM_CODIGO = " & Chr$(34) + Text1.text + Chr$(34)
         Data1.Recordset.FindFirst criterio
         If Not Data1.Recordset.NoMatch Then
            DBGrid1.Visible = False
            MsgBox "La Familia ya ha sido registrada !"
            DBGrid1.Visible = False
            Text1.text = ""
            Text1.SetFocus
         Else
            Text1.Enabled = False
            Text1.Visible = True
            Text2.Visible = True
            Text2.SetFocus
         End If
      End If
    End If
    Label1.Visible = True
    
End Sub





