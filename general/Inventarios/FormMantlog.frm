VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form FormMantProv 
   Caption         =   "Mantenimiento de Provedores"
   ClientHeight    =   5820
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9180
   LinkTopic       =   "Form11"
   ScaleHeight     =   5820
   ScaleWidth      =   9180
   StartUpPosition =   3  'Windows Default
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   120
      Top             =   4560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Inventarios.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   7800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "MaeProv"
      Top             =   360
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton Command20 
      Caption         =   "Direcciones"
      Height          =   855
      Left            =   4200
      Picture         =   "FormMantProv.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   4800
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command18 
      Caption         =   "&Cancelar B�squeda"
      Height          =   855
      Left            =   6600
      Picture         =   "FormMantProv.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   4800
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command17 
      Caption         =   "Buscar por &Rz. Social"
      Height          =   855
      Left            =   4200
      MaskColor       =   &H00C0FFFF&
      Picture         =   "FormMantProv.frx":058C
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   4800
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command16 
      Caption         =   "&Buscar por &RUC"
      Height          =   855
      Left            =   1800
      MaskColor       =   &H00C0FFFF&
      Picture         =   "FormMantProv.frx":09CE
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   4800
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command15 
      Caption         =   "&Cancelar Orden"
      Height          =   855
      Left            =   6600
      Picture         =   "FormMantProv.frx":0E10
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   4800
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Por &Raz�n Social"
      Height          =   855
      Left            =   4200
      Picture         =   "FormMantProv.frx":0F5A
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   4800
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Por &C�digo"
      Height          =   855
      Left            =   1800
      Picture         =   "FormMantProv.frx":10A4
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   4800
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command6 
      Caption         =   "&Ordenar "
      Height          =   855
      Left            =   6600
      Picture         =   "FormMantProv.frx":11EE
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   4800
      Width           =   975
   End
   Begin VB.CommandButton Command12 
      Caption         =   "C&ancelar Imprimir"
      Height          =   855
      Left            =   6600
      Picture         =   "FormMantProv.frx":14F8
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   4800
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Por &Raz�n Social"
      Height          =   855
      Left            =   4200
      Picture         =   "FormMantProv.frx":1642
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   4800
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Por &C�digo"
      Height          =   855
      Left            =   1800
      Picture         =   "FormMantProv.frx":194C
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4800
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command9 
      Caption         =   "&Cancelar"
      Height          =   855
      Left            =   6600
      Picture         =   "FormMantProv.frx":1C56
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4800
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command8 
      Caption         =   "&Grabar"
      Height          =   855
      Left            =   1800
      Picture         =   "FormMantProv.frx":1DA0
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4800
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text9 
      Height          =   285
      Left            =   5040
      TabIndex        =   12
      Top             =   2640
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   2040
      TabIndex        =   11
      Top             =   3240
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   2040
      TabIndex        =   10
      Top             =   2640
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   2040
      TabIndex        =   9
      Top             =   2040
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   2040
      TabIndex        =   8
      Top             =   1440
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2040
      TabIndex        =   7
      Top             =   840
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2040
      TabIndex        =   6
      Top             =   240
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton Command7 
      Caption         =   "&Salir"
      Height          =   855
      Left            =   7800
      Picture         =   "FormMantProv.frx":20AA
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4800
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   5400
      Picture         =   "FormMantProv.frx":21F4
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4800
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Buscar"
      Height          =   855
      Left            =   4200
      Picture         =   "FormMantProv.frx":2636
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4800
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Eliminar"
      Height          =   855
      Left            =   3000
      Picture         =   "FormMantProv.frx":2A78
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4800
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Modificar"
      Height          =   855
      Left            =   1800
      Picture         =   "FormMantProv.frx":2EBA
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4800
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Adicionar"
      Height          =   855
      Left            =   600
      Picture         =   "FormMantProv.frx":32FC
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4800
      Width           =   975
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "FormMantProv.frx":3606
      Height          =   4575
      Left            =   240
      OleObjectBlob   =   "FormMantProv.frx":361A
      TabIndex        =   26
      Top             =   120
      Width           =   8775
   End
   Begin VB.Label Label10 
      Height          =   255
      Left            =   4200
      TabIndex        =   34
      Top             =   2040
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label Label9 
      Caption         =   "Fax:"
      Height          =   255
      Left            =   4080
      TabIndex        =   33
      Top             =   2640
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label6 
      Caption         =   "R.U.C."
      Height          =   255
      Left            =   480
      TabIndex        =   32
      Top             =   3240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Tel�fono"
      Height          =   255
      Left            =   480
      TabIndex        =   31
      Top             =   2640
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Zona Postal:"
      Height          =   255
      Left            =   480
      TabIndex        =   30
      Top             =   2040
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Direcci�n"
      Height          =   255
      Left            =   480
      TabIndex        =   29
      Top             =   1440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Nombre"
      Height          =   255
      Left            =   480
      TabIndex        =   28
      Top             =   840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Codigo"
      Height          =   255
      Left            =   480
      TabIndex        =   27
      Top             =   240
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "FormMantProv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim resp As String
Private Sub Command1_Click()
   resp = "S"
   
   Label10.Caption = ""
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
  Dim CODIGO2 As String
  
'   Data1.RecordSource = "SELECT * FROM MAEPROV ORDER BY PRVCCODIGO"
'   Data1.Refresh
   CrystalReport1.ReportFileName = "C:\catprov.rpt"
   CrystalReport1.DiscardSavedData = True
   CrystalReport1.Formulas(0) = "EMPRESA='" & CODIGO2 & "'"
   CrystalReport1.Destination = crptToWindow
   CrystalReport1.Action = 1
End Sub

Private Sub Command11_Click()
  Dim CODIGO2 As String

   CrystalReport1.ReportFileName = "C:\catprovdes.rpt"
   CrystalReport1.DiscardSavedData = True
   CrystalReport1.Formulas(0) = "EMPRESA='" & CODIGO2 & "'"
   CrystalReport1.Destination = crptToWindow
   CrystalReport1.Action = 1
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
End Sub

Private Sub Command16_Click()
 
DBGrid1.Enabled = True
DBGrid1.Visible = True

src = InputBox$("Ingrese RUC", "B�squeda")
ncar = Str$(Len(src)) 'PRVCRUC
criterio = "MID$(PRVCRUC,1," + ncar + ") = " & Chr$(34) + src + Chr$(34)
Data1.Recordset.FindFirst criterio
If Data1.Recordset.NoMatch Then
   MsgBox "No se encontr� el Registro !"
End If



End Sub

Private Sub Command17_Click()
DBGrid1.Enabled = True
DBGrid1.Visible = True

src = InputBox$("Ingrese Nombre de la empresa", "B�squeda")
   ncar = Str$(Len(src))
   criterio = "MID$(PRVCNOMBRE,1," + ncar + ") = " & Chr$(34) + src + Chr$(34)
   Data1.Recordset.FindFirst criterio
   If Data1.Recordset.NoMatch Then
      MsgBox "No se encontr� el Registro !"
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
   
   
   
  End Sub

Private Sub Command19_Click()
   Form3.Show
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
   
   Command20.Visible = True
    
   
        
  
   Label1.Visible = True
   Label2.Visible = True
   Label3.Visible = True
   Label4.Visible = True
   Label5.Visible = True
   Label6.Visible = True
 
   
   Label9.Visible = True
   
   
   
   
   
   Text1.text = Data1.Recordset("PRVCCODIGO")
   
   If Not IsNull(Data1.Recordset("PRVCNOMBRE")) Then
      Text2.text = Data1.Recordset("PRVCNOMBRE")
   Else
      Text2.text = ""
    End If
   
   If Not IsNull(Data1.Recordset("PRVCDIRECC")) Then
      Text3.text = Data1.Recordset("PRVCDIRECC")
   Else
      Text3.text = ""
   End If
   
   
   'End If
   If Not IsNull(Data1.Recordset("PRVCTELEF1")) Then
      Text5.text = Data1.Recordset("PRVCTELEF1")
   Else
      Text5.text = ""
   End If
   
   If Not IsNull(Data1.Recordset("PRVCRUC")) Then
      Text6.text = Data1.Recordset("PRVCRUC")
   Else
      Text6.text = ""
   End If
    
   
   If Not IsNull(Data1.Recordset("PRVCFAX")) Then
      Text9.text = Data1.Recordset("PRVCFAX")
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
  '   Form4.Show
End Sub

Private Sub Command3_Click()
   
  Dim MENSAJE$
  
  DBGrid1.Enabled = True
  DBGrid1.Visible = True
   
 
  Command1.Visible = True
  Command2.Visible = True
  Command3.Visible = True
  Command4.Visible = True
  Command5.Visible = True
  Command6.Visible = True
  Command7.Visible = True
 
    
     
   MENSAJE$ = "Est� seguro de Borrar ?" & Chr$(10)
   MENSAJE$ = MENSAJE$ & Data1.Recordset(1)
   If MsgBox(MENSAJE$, 33, "Borra la empresa ?") <> 1 Then Exit Sub
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
Command5.Visible = False
Command6.Visible = False
Command7.Visible = False

Command16.Visible = True
Command17.Visible = True
Command18.Visible = True


 





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
   
End Sub

Private Sub Command13_Click()
   Data1.RecordSource = "Select * from MAEPROV order by PRVCRUC"
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
   Command5.Visible = True
   Command6.Visible = True
   Command7.Visible = True
   Command8.Visible = False
   Command20.Visible = False
   

   
      
   If Text1.text <> "" Then
      If resp = "S" Then
         Data1.Recordset.AddNew
         Data1.Recordset("PRVCCODIGO") = UCase$(Text1.text)
      Else
         Data1.Recordset.Edit
      End If
      
      
      If Text2.text <> "" Then
         Data1.Recordset("PRVCNOMBRE") = Mid$(UCase$(Text2.text), 1, 50)
      Else
         Data1.Recordset("PRVCNOMBRE") = " "
      End If
      
      If Text3.text <> "" Then
         Data1.Recordset("PRVCDIRECC") = Mid$(UCase$(Text3.text), 1, 50)
      Else
         Data1.Recordset("PRVCDIRECC") = " "
      End If
      
      If Text4.text <> "" Then
         'Data1.Recordset("PRVCDIRECC") = Mid$(UCase$(Text4.Text), 1, 8)
      Else
         'Data1.Recordset("PRVCDIRECC") = " "
      End If
      
      
      If Text5.text <> "" Then
         Data1.Recordset("PRVCTELEF1") = Mid$(UCase$(Text5.text), 1, 25)
      Else
         Data1.Recordset("PRVCTELEF1") = " "
      End If
     
      
      If Text6.text <> "" Then
         Data1.Recordset("PRVCRUC") = Mid$(UCase$(Text6.text), 1, 8)
         Else
         Data1.Recordset("PRVCRUC") = " "
         
      End If
      
      
        
      
      
      If Text9.text <> "" Then
         'Data1.Recordset("PRVCFAX") = Mid$(UCase$(Text9.Text), 1, 15)
      Else
         'Data1.Recordset("PRVCFAX") = " "
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
   Command5.Visible = True
   Command6.Visible = True
   Command7.Visible = True
   
   Command13.Visible = False
   Command14.Visible = False
   Command15.Visible = False
 
  
End Sub

Private Sub Command14_Click()
   Data1.RecordSource = "Select * from MAEPROV order by PRVCNOMBRE"
   Data1.Refresh
End Sub

Private Sub Command9_Click()
   DBGrid1.Visible = True
   
   Command1.Visible = True
   Command2.Visible = True
   Command3.Visible = True
   Command4.Visible = True
   Command5.Visible = True
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
Init_ControlDataGrid DBGrid1
'Init_ControlDataGrid (DBGrid1)
central FormMantProv


End Sub

Private Sub Text1_Change()
    
    If resp = "S" Then
      If Len(Text1.text) = 8 Then
         criterio = "PRVCRUC = " & Chr$(34) + Text1.text + Chr$(34)
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

Private Sub Text4_Change()
   
   If Len(Text4.text) = 6 Then
      criterio = "PRVCCODIGO = " & Chr$(34) + Text4.text + Chr$(34)
      Data2.Recordset.FindFirst criterio
      If Data2.Recordset.NoMatch Then
         MsgBox "El C�digo Postal no ha sido registrado !"
         Text4.text = ""
      Else
         Text4.text = UCase$(Text4.text)
         Label10.Caption = Data2.Recordset(1)
         Label10.Visible = True
      End If
   End If
End Sub



