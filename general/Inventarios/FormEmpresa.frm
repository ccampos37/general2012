VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form FormEmpresa 
   Caption         =   "Mantenimiento de Empresa"
   ClientHeight    =   5745
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7845
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   ScaleHeight     =   5745
   ScaleWidth      =   7845
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&Adicionar"
      Height          =   675
      Left            =   480
      Picture         =   "FormEmpresa.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   4920
      Width           =   775
   End
   Begin VB.CommandButton Command7 
      Caption         =   "&Salir"
      Height          =   675
      Left            =   6480
      Picture         =   "FormEmpresa.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   4920
      Width           =   775
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1920
      MaxLength       =   11
      TabIndex        =   20
      Top             =   360
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1920
      TabIndex        =   19
      Top             =   840
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1935
      TabIndex        =   18
      Top             =   1290
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   1920
      TabIndex        =   17
      Top             =   1770
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   1920
      TabIndex        =   16
      Top             =   2760
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   1920
      TabIndex        =   15
      Top             =   3720
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox Text9 
      Height          =   285
      Left            =   1920
      TabIndex        =   14
      Top             =   3240
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton Command8 
      Caption         =   "&Grabar"
      Height          =   675
      Left            =   1680
      Picture         =   "FormEmpresa.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4920
      Visible         =   0   'False
      Width           =   775
   End
   Begin VB.CommandButton Command9 
      Caption         =   "&Cancelar"
      Height          =   675
      Left            =   5280
      Picture         =   "FormEmpresa.frx":0A56
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4920
      Visible         =   0   'False
      Width           =   775
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Por &Código"
      Height          =   675
      Left            =   1680
      Picture         =   "FormEmpresa.frx":0BA0
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4920
      Visible         =   0   'False
      Width           =   775
   End
   Begin VB.CommandButton Command12 
      Caption         =   "C&ancelar Imprimir"
      Height          =   675
      Left            =   5280
      Picture         =   "FormEmpresa.frx":0EAA
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4920
      Visible         =   0   'False
      Width           =   775
   End
   Begin VB.CommandButton Command6 
      Caption         =   "&Ordenar "
      Height          =   675
      Left            =   5280
      Picture         =   "FormEmpresa.frx":0FF4
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4920
      Width           =   775
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Por &Código"
      Height          =   675
      Left            =   1680
      Picture         =   "FormEmpresa.frx":12FE
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4920
      Visible         =   0   'False
      Width           =   775
   End
   Begin VB.CommandButton Command15 
      Caption         =   "&Cancelar Orden"
      Height          =   675
      Left            =   5280
      Picture         =   "FormEmpresa.frx":1448
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4920
      Visible         =   0   'False
      Width           =   775
   End
   Begin VB.CommandButton Command16 
      Caption         =   "&Buscar por Codigo"
      Height          =   675
      Left            =   1680
      MaskColor       =   &H00C0FFFF&
      Picture         =   "FormEmpresa.frx":1592
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4920
      Visible         =   0   'False
      Width           =   775
   End
   Begin VB.CommandButton Command18 
      Caption         =   "&Cancelar Búsqueda"
      Height          =   675
      Left            =   5280
      Picture         =   "FormEmpresa.frx":19D4
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4920
      Visible         =   0   'False
      Width           =   775
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Inventarios.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   7320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "TablaEmp"
      Top             =   120
      Visible         =   0   'False
      Width           =   1140
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "FormEmpresa.frx":1B1E
      Height          =   4575
      Left            =   195
      OleObjectBlob   =   "FormEmpresa.frx":1B32
      TabIndex        =   13
      Top             =   210
      Width           =   7455
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   1920
      TabIndex        =   34
      Top             =   2265
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Eliminar"
      Height          =   675
      Left            =   2880
      Picture         =   "FormEmpresa.frx":286D
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   4920
      Width           =   775
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Buscar"
      Height          =   675
      Left            =   4080
      Picture         =   "FormEmpresa.frx":2CAF
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   4920
      Width           =   775
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Por &Razón Social"
      Height          =   675
      Left            =   4080
      Picture         =   "FormEmpresa.frx":30F1
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4920
      Visible         =   0   'False
      Width           =   775
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Por &Razón Social"
      Height          =   675
      Left            =   4080
      Picture         =   "FormEmpresa.frx":33FB
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4920
      Visible         =   0   'False
      Width           =   775
   End
   Begin VB.CommandButton Command17 
      Caption         =   "Buscar por &Nombre"
      Height          =   675
      Left            =   4080
      MaskColor       =   &H00C0FFFF&
      Picture         =   "FormEmpresa.frx":3545
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4920
      Visible         =   0   'False
      Width           =   775
   End
   Begin VB.CommandButton Command20 
      Caption         =   "Direcciones"
      Height          =   675
      Left            =   4080
      Picture         =   "FormEmpresa.frx":3987
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4920
      Visible         =   0   'False
      Width           =   775
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Modificar"
      Height          =   675
      Left            =   1680
      Picture         =   "FormEmpresa.frx":3DC9
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   4920
      Width           =   775
   End
   Begin VB.Label Label7 
      Caption         =   "Departamento"
      Height          =   255
      Left            =   360
      TabIndex        =   33
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "R.U.C:"
      Height          =   255
      Left            =   360
      TabIndex        =   32
      Top             =   360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Nombre"
      Height          =   255
      Left            =   360
      TabIndex        =   31
      Top             =   840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Dirección"
      Height          =   255
      Left            =   360
      TabIndex        =   30
      Top             =   1320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Distrito"
      Height          =   255
      Left            =   360
      TabIndex        =   29
      Top             =   1800
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Teléfono"
      Height          =   255
      Left            =   360
      TabIndex        =   28
      Top             =   2760
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Rubro:"
      Height          =   255
      Left            =   360
      TabIndex        =   27
      Top             =   3720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label9 
      Caption         =   "Fax:"
      Height          =   255
      Left            =   360
      TabIndex        =   26
      Top             =   3240
      Visible         =   0   'False
      Width           =   375
   End
End
Attribute VB_Name = "FormEmpresa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim resp As String
Private Sub Command1_Click()
   resp = "S"
   
  
   DBGrid1.Visible = False
   Command1.Visible = False
   Command2.Visible = False
   Command3.Visible = False
   Command4.Visible = False
   
   Command6.Visible = False
   Command7.Visible = False
   
   Command8.Visible = True
   Command9.Visible = True
   
   Label1.Visible = True
   Label2.Visible = True
   Label3.Visible = True
   Label4.Visible = True
   Label5.Visible = True
   Label6.Visible = True
  
   
   Label9.Visible = True
   
  
   Text1.Visible = True
   Text2.Visible = True
   Text3.Visible = True
   Text4.Visible = True
   Text5.Visible = True
   Text6.Visible = True
  
  
   Text9.Visible = True
      
Text1.SetFocus
End Sub

Private Sub Command10_Click()
   Data1.RecordSource = "Select * from EMPRESA order by EMP_CODIGO"
   Data1.Refresh
End Sub

Private Sub Command11_Click()
   Data1.RecordSource = "Select * from EMPRESA order by EMP_RAZON_NOMBRE"
   Data1.Refresh
End Sub

Private Sub Command12_Click()
   DBGrid1.Visible = True
   Command1.Visible = True
   Command2.Visible = True
   Command3.Visible = True
   Command4.Visible = True
   'Command5.Visible = True
   Command6.Visible = True
   Command7.Visible = True
   Command10.Visible = False
   Command11.Visible = False
End Sub

Private Sub Command16_Click()
 
DBGrid1.Enabled = True
DBGrid1.Visible = True

src = InputBox$("Ingrese RUC", "Búsqueda")
ncar = Str$(Len(src))
criterio = "left(ERUC," & ncar & ") = '" & src & "'"
Data1.Recordset.FindFirst criterio
If Data1.Recordset.NoMatch Then
   MsgBox "No se encontró el Registro !"
End If



End Sub

Private Sub Command17_Click()
DBGrid1.Enabled = True
DBGrid1.Visible = True

  src = InputBox$("Ingrese Nombre de la empresa", "Búsqueda")
   ncar = Str$(Len(src))
   criterio = "Left(ENOMBRE," & ncar & ") = '" & src & "'"
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
   'Command5.Visible = True
   Command6.Visible = True
   Command7.Visible = True
   Command16.Visible = False
   Command17.Visible = False
   Command18.Visible = False
   
   
   
  End Sub

Private Sub Command19_Click()
   Form3.Show 1
End Sub

Private Sub Command2_Click()
   resp = "N"
   DBGrid1.Visible = False
   
   Command1.Visible = False
   Command2.Visible = False
   Command3.Visible = False
   Command4.Visible = False
   'Command5.Visible = False
   Command6.Visible = False
   Command7.Visible = False
   
   Command8.Visible = True
   Command9.Visible = True
   
   Command20.Visible = True
    
   
        
  
   Label1.Visible = True
   Label2.Visible = True
   Label3.Visible = True
   Label4.Visible = True
   Label5.Visible = True
   Label6.Visible = True
 
   
   Label9.Visible = True
   
   
   
   
   
   Text1.text = Data1.Recordset("ECODIGO")
   
   If Not IsNull(Data1.Recordset("ENOMBRE")) Then
      Text2.text = Data1.Recordset("ENOMBRE")
      VGNemp = Mid(Text2, 1, 28)
   Else
      Text2.text = ""
    End If
   
   If Not IsNull(Data1.Recordset("EDIREC")) Then
      Text3.text = Data1.Recordset("EDIREC")
   Else
      Text3.text = ""
   End If
   
   
   'End If
   If Not IsNull(Data1.Recordset("ETELEF")) Then
      Text5.text = Data1.Recordset("ETELEF")
   Else
      Text5.text = ""
   End If
   
   If Not IsNull(Data1.Recordset("ERUBRO")) Then
      Text6.text = Data1.Recordset("ERUBRO")
   Else
      Text6.text = ""
   End If
    
   
   If Not IsNull(Data1.Recordset("EFAX")) Then
      Text9.text = Data1.Recordset("EFAX")
   Else
      Text9.text = ""
   End If
   
   
   Text1.Visible = True
   Text1.Enabled = False
   Text2.Visible = True
   Text3.Visible = True
   Text4.Visible = True
   Text5.Visible = True
   Text6.Visible = True
 
   Text9.Visible = True
   
   
   
  
   Text2.SetFocus
End Sub

Private Sub Command20_Click()
  '   Form4.show 1
End Sub

Private Sub Command3_Click()
   
  Dim Mensaje$
  
  DBGrid1.Enabled = True
  DBGrid1.Visible = True
   
 
  Command1.Visible = True
  Command2.Visible = True
  Command3.Visible = True
  Command4.Visible = True
  'Command5.Visible = True
  Command6.Visible = True
  Command7.Visible = True
 
    
     
   Mensaje$ = "Está seguro de Borrar ?" & Chr$(10)
   Mensaje$ = Mensaje$ & Data1.Recordset(1)
   If MsgBox(Mensaje$, 33, "Borra la empresa ?") <> 1 Then Exit Sub
   Data1.Recordset.Delete
   Data1.Recordset.MoveNext
   If Data1.Recordset.EOF Then Data1.Recordset.MovePrevious



   
 

 
End Sub

Private Sub Command4_Click()
DBGrid1.Visible = True
DBGrid1.Enabled = True



Command1.Visible = False
Command2.Visible = False
Command3.Visible = False
Command4.Visible = False
'Command5.Visible = False
Command6.Visible = False
Command7.Visible = False

Command16.Visible = True
Command17.Visible = True
Command18.Visible = True


 





End Sub

'Private Sub Command5_Click()
'
'   Command1.Visible = False
'   Command2.Visible = False
'   Command3.Visible = False
'   Command4.Visible = False
'   Command5.Visible = False
'   Command6.Visible = False
'   Command7.Visible = False
'
'   Command10.Visible = True
'   Command11.Visible = True
'   Command12.Visible = True
'
'
'
'
'
'
'End Sub

Private Sub Command6_Click()
  
   
   Command1.Visible = False
   Command2.Visible = False
   Command3.Visible = False
   Command4.Visible = False
   Command6.Visible = False
   Command7.Visible = False
     
   Command13.Visible = True
   Command14.Visible = True
   Command15.Visible = True
   
End Sub

Private Sub Command13_Click()
   Data1.RecordSource = "Select * from Empresa order by EMP_RUC_DOCUMENTO"
   Data1.Refresh
End Sub
Private Sub Command7_Click()
   Unload Me
End Sub

Private Sub Command8_Click()

   Command1.Visible = True
   Command2.Visible = True
   Command3.Visible = True
   Command4.Visible = True
   Command6.Visible = True
   Command7.Visible = True
   Command8.Visible = False
   Command20.Visible = False
   

   
      
   If Not IsNull(Text1.text) And Text1.text <> "" Then
      If resp = "S" Then
         Data1.Recordset.AddNew
         Data1.Recordset("ECODIGO") = Format(Data1.Recordset.RecordCount, "000")
         Data1.Recordset("ERUC") = UCase$(Text1.text)
      Else
         Data1.Recordset.Edit
      End If
      
      
      If Not IsNull(Text2.text) Then
         Data1.Recordset("ENOMBRE") = Mid$(UCase$(Text2.text), 1, 50)
      Else
         Data1.Recordset("ENOMBRE") = " "
      End If
      
      If Not IsNull(Text3.text) Then
         Data1.Recordset("EDIREC") = Mid$(UCase$(Text3.text), 1, 50)
      Else
         Data1.Recordset("EDIREC") = " "
      End If
      
      If Not IsNull(Text4.text) Then
         'Data1.Recordset("EDIREC") = Mid$(UCase$(Text4.Text), 1, 8)
      Else
         'Data1.Recordset("EDIREC") = " "
      End If
      
      
      If Not IsNull(Text5.text) Then
         Data1.Recordset("ETELEF") = Mid$(UCase$(Text5.text), 1, 25)
      Else
         Data1.Recordset("ETELEF") = " "
      End If
     
      
        If Text6.text <> "" Then
         Data1.Recordset("ERUBRO") = Mid$(UCase$(Text6.text), 1, 2)
         Else
         Data1.Recordset("ERUBRO") = " "
         
      End If
       If Not IsNull(Text9.text) Then
         'Data1.Recordset("EFAX") = Mid$(UCase$(Text9.Text), 1, 15)
      Else
         'Data1.Recordset("EFAX") = " "
      End If
      
      
    End If
            
   
   Data1.Recordset.Update
   Data1.Refresh
   
   
   Label1.Visible = False
   Label2.Visible = False
   Label3.Visible = False
   Label4.Visible = False
   Label5.Visible = False
   Label6.Visible = False
 
   Label9.Visible = False
   
   
   Text1.Enabled = True
   
   Text1.Visible = False
   Text2.Visible = False
   Text3.Visible = False
   Text4.Visible = False
   Text5.Visible = False
   Text6.Visible = False
 
   Text9.Visible = False
   
   
   
   Text1 = ""
   Text2 = ""
   Text3 = ""
   Text4 = ""
   Text5 = ""
   Text6 = ""
   Text7 = ""
   Text8 = ""
   Text9 = ""
   
  DBGrid1.Visible = True
 DBGrid1.Enabled = True
 
 
End Sub

Private Sub Command15_Click()
   Command1.Visible = True
   Command2.Visible = True
   Command3.Visible = True
   Command4.Visible = True
   Command6.Visible = True
   Command7.Visible = True
   
   Command13.Visible = False
   Command14.Visible = False
   Command15.Visible = False
 
  
End Sub

Private Sub Command14_Click()
   Data1.RecordSource = "Select * from EMPRESA order by EMP_RAZON_NOMBRE"
   Data1.Refresh
End Sub

Private Sub Command9_Click()
   DBGrid1.Visible = True
   
   Command1.Visible = True
   Command2.Visible = True
   Command3.Visible = True
   Command4.Visible = True
   Command6.Visible = True
   Command7.Visible = True
   
   Command8.Visible = False
   Command9.Visible = False
   Command20.Visible = False
      
   Label1.Visible = False
   Label2.Visible = False
   Label3.Visible = False
   Label4.Visible = False
   Label5.Visible = False
   Label6.Visible = False
  
   Label9.Visible = False
   
   Text1.Enabled = True
   Text1.Visible = False
   Text2.Visible = False
   Text3.Visible = False
   Text4.Visible = False
   Text5.Visible = False
   Text6.Visible = False
  
   Text9.Visible = False
   
   Text1 = ""
   Text2 = ""
   Text3 = ""
   Text4 = ""
   Text5 = ""
   Text6 = ""
   Text7 = ""
   Text8 = ""
   Text9 = ""
End Sub

Private Sub Form_Load()
Label1.Enabled = True
Label2.Enabled = True
Label3.Enabled = True
Label4.Enabled = True
Label5.Enabled = True
Label6.Enabled = True


Label1.Visible = False
Label2.Visible = False
Label3.Visible = False
Label4.Visible = False
Label5.Visible = False
Label6.Visible = False
Data1.DatabaseName = cRuta2
Init_ControlDBGrid DBGrid1
central FormEmpresa


End Sub

Private Sub Text1_Change()
    
    If resp = "S" Then
      If Len(Text1.text) = 8 Then
         criterio = "ERUC = '" & Text1.text & "'"
         Data1.Recordset.FindFirst criterio
         If Not Data1.Recordset.NoMatch Then
            DBGrid1.Visible = False
            MsgBox "El Tercero ya ha sido registrado !"
            DBGrid1.Visible = False
            
            Text1.text = ""
            
            Text1.SetFocus
         Else
            Text1.Enabled = False
            
            Text1.Visible = True
            
            Text2.Visible = True
            Text3.Visible = True
            Text4.Visible = True
            Text5.Visible = True
            Text6.Visible = True
            Text7.Visible = True
            
            Text8.Visible = True
            Text9.Visible = True
            Text2.SetFocus
         
         End If
       End If
   End If
         
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 And Text1 <> "" Then
     SendKeys "{tab}"
     KeyAscii = 0
  End If
     
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 And Text2 <> "" Then
     SendKeys "{tab}"
     KeyAscii = 0
  End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
     SendKeys "{tab}"
     KeyAscii = 0
  End If
End Sub

Private Sub Text4_Change()
   
   If Len(Text4.text) = 6 Then
      criterio = "ECODIGO = '" & Text4.text & "'"
      Data2.Recordset.FindFirst criterio
      If Data2.Recordset.NoMatch Then
         MsgBox "El Código Postal no ha sido registrado !"
         Text4.text = ""
      Else
         Text4.text = UCase$(Text4.text)
'         Label10.Caption = Data2.Recordset(1)
'         Label10.Visible = True
      End If
   End If
End Sub


