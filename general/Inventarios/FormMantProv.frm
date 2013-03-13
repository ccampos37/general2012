VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form FormMantProv 
   Caption         =   "Provedores"
   ClientHeight    =   5820
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9180
   LinkTopic       =   "Form11"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   9180
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&Adicionar"
      Height          =   775
      Left            =   600
      Picture         =   "FormMantProv.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   4800
      Width           =   840
   End
   Begin VB.Frame Frame1 
      Height          =   4215
      Left            =   360
      TabIndex        =   30
      Top             =   240
      Visible         =   0   'False
      Width           =   8175
      Begin VB.TextBox Text12 
         Height          =   285
         Left            =   5040
         TabIndex        =   11
         Top             =   3600
         Width           =   1575
      End
      Begin VB.TextBox Text11 
         Height          =   285
         Left            =   2040
         TabIndex        =   10
         Top             =   3600
         Width           =   1695
      End
      Begin VB.TextBox Text10 
         Height          =   285
         Left            =   2040
         TabIndex        =   9
         Top             =   3240
         Width           =   1695
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   2040
         TabIndex        =   6
         Top             =   2520
         Width           =   375
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   2040
         TabIndex        =   5
         Top             =   2160
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2040
         TabIndex        =   0
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2040
         TabIndex        =   1
         Top             =   720
         Width           =   5415
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   2040
         TabIndex        =   3
         Top             =   1440
         Width           =   5415
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   2040
         TabIndex        =   4
         Top             =   1800
         Width           =   1695
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   2040
         TabIndex        =   7
         Top             =   2880
         Width           =   1695
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   2040
         MaxLength       =   8
         TabIndex        =   2
         Top             =   1080
         Width           =   1695
      End
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   5040
         TabIndex        =   8
         Top             =   2880
         Width           =   1575
      End
      Begin VB.Label Label14 
         Caption         =   "Telf Repres."
         Height          =   255
         Left            =   3960
         TabIndex        =   43
         Top             =   3600
         Width           =   975
      End
      Begin VB.Label Label12 
         Caption         =   "Cargo Repres."
         Height          =   255
         Left            =   480
         TabIndex        =   42
         Top             =   3600
         Width           =   1095
      End
      Begin VB.Label Label11 
         Caption         =   "Representante"
         Height          =   255
         Left            =   480
         TabIndex        =   41
         Top             =   3240
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "Giro del Proveedor"
         Height          =   255
         Left            =   480
         TabIndex        =   40
         Top             =   2520
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "Pais de Procedencia"
         Height          =   255
         Left            =   480
         TabIndex        =   39
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Codigo"
         Height          =   255
         Left            =   480
         TabIndex        =   38
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Nombre"
         Height          =   255
         Left            =   480
         TabIndex        =   37
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Dirección"
         Height          =   255
         Left            =   480
         TabIndex        =   36
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Localidad"
         Height          =   255
         Left            =   480
         TabIndex        =   35
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Teléfono"
         Height          =   255
         Left            =   480
         TabIndex        =   34
         Top             =   2880
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "R.U.C."
         Height          =   255
         Left            =   480
         TabIndex        =   33
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "Fax:"
         Height          =   255
         Left            =   4440
         TabIndex        =   32
         Top             =   2880
         Width           =   495
      End
      Begin VB.Label Label10 
         Height          =   255
         Left            =   4080
         TabIndex        =   31
         Top             =   2280
         Visible         =   0   'False
         Width           =   2295
      End
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   0
      Top             =   4800
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
   Begin VB.CommandButton Command18 
      Caption         =   "&Cancelar Búsqueda"
      Height          =   775
      Left            =   6600
      Picture         =   "FormMantProv.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   4800
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.CommandButton Command17 
      Caption         =   "Por &Rz. Social"
      Height          =   775
      Left            =   4200
      MaskColor       =   &H00C0FFFF&
      Picture         =   "FormMantProv.frx":0454
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   4800
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.CommandButton Command16 
      Caption         =   "Por &Codigo"
      Height          =   775
      Left            =   1680
      MaskColor       =   &H00C0FFFF&
      Picture         =   "FormMantProv.frx":0896
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   4320
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.CommandButton Command15 
      Caption         =   "&Cancelar Orden"
      Height          =   775
      Left            =   6600
      Picture         =   "FormMantProv.frx":0CD8
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   4800
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Por &Razón Social"
      Height          =   775
      Left            =   4200
      Picture         =   "FormMantProv.frx":0E22
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   4800
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Por &Código"
      Height          =   775
      Left            =   1800
      Picture         =   "FormMantProv.frx":0F6C
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   4800
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.CommandButton Command6 
      Caption         =   "&Ordenar "
      Height          =   775
      Left            =   6600
      Picture         =   "FormMantProv.frx":10B6
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   4800
      Width           =   840
   End
   Begin VB.CommandButton Command12 
      Caption         =   "C&ancelar Imprimir"
      Height          =   775
      Left            =   6600
      Picture         =   "FormMantProv.frx":13C0
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   4800
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Por &Razón Social"
      Height          =   775
      Left            =   4200
      Picture         =   "FormMantProv.frx":150A
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   4800
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Por &Código"
      Height          =   775
      Left            =   1800
      Picture         =   "FormMantProv.frx":1814
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   4800
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.CommandButton Command9 
      Caption         =   "&Cancelar"
      Height          =   775
      Left            =   6600
      Picture         =   "FormMantProv.frx":1B1E
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   4800
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.CommandButton Command8 
      Caption         =   "&Grabar"
      Height          =   775
      Left            =   1800
      Picture         =   "FormMantProv.frx":1C68
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   4800
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.CommandButton Command7 
      Caption         =   "&Salir"
      Height          =   775
      Left            =   7800
      Picture         =   "FormMantProv.frx":20AA
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   4800
      Width           =   840
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&Imprimir"
      Height          =   775
      Left            =   5400
      Picture         =   "FormMantProv.frx":24EC
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4800
      Width           =   840
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Buscar"
      Height          =   775
      Left            =   4200
      Picture         =   "FormMantProv.frx":292E
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4800
      Width           =   840
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Eliminar"
      Height          =   775
      Left            =   3000
      Picture         =   "FormMantProv.frx":2D70
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4800
      Width           =   840
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Modificar"
      Height          =   775
      Left            =   1800
      Picture         =   "FormMantProv.frx":31B2
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4800
      Width           =   840
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "FormMantProv.frx":35F4
      Height          =   4575
      Left            =   240
      OleObjectBlob   =   "FormMantProv.frx":3608
      TabIndex        =   29
      Top             =   120
      Width           =   8775
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
   Frame1.Visible = True
  Text1.Enabled = True
   Text1.SetFocus
End Sub



Private Sub Command10_Click()
  Dim CODIGO2 As String
  
'   Data1.RecordSource = "SELECT * FROM MAEPROV ORDER BY PRVCCODIGO"
'   Data1.Refresh
   CrystalReport1.ReportFileName = RUTA & "REPORTEINV\catprov.rpt"
   Ubi_Tab CrystalReport1
   CrystalReport1.DiscardSavedData = True
   CrystalReport1.Formulas(0) = "EMPRESA='" & CODIGO2 & "'"
   CrystalReport1.Destination = crptToWindow
   If CrystalReport1.Status <> 2 Then
      CrystalReport1.Action = 1
   End If
End Sub

Private Sub Command11_Click()
  Dim CODIGO2 As String

   CrystalReport1.ReportFileName = RUTA & "REPORTEINV\catprovdes.rpt"
   Ubi_Tab CrystalReport1
   CrystalReport1.DiscardSavedData = True
   CrystalReport1.Formulas(0) = "EMPRESA='" & CODIGO2 & "'"
   CrystalReport1.Destination = crptToWindow
     If CrystalReport1.Status <> 2 Then
      CrystalReport1.Action = 1
   End If
End Sub


Private Sub Command12_Click()
  Frame1.Visible = False
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
Dim criterio As String
 Frame1.Visible = False
DBGrid1.Enabled = True
DBGrid1.Visible = True

src = InputBox$("Ingrese Codigo", "Búsqueda")
ncar = Str$(Len(src)) 'PRVCRUC
criterio = "MID$(PRVCcodigo,1," + ncar + ") = " & Chr$(34) + src + Chr$(34)
Data1.Recordset.FindFirst criterio
If Data1.Recordset.NoMatch Then
   MsgBox "No se encontró el Registro !"
End If



End Sub

Private Sub Command17_Click()
 Frame1.Visible = False
DBGrid1.Enabled = True
DBGrid1.Visible = True

src = InputBox$("Ingrese Nombre de la empresa", "Búsqueda")
   ncar = Str$(Len(src))
   criterio = "MID$(PRVCNOMBRE,1," + ncar + ") = " & Chr$(34) + src + Chr$(34)
   Data1.Recordset.FindFirst criterio
   If Data1.Recordset.NoMatch Then
      MsgBox "No se encontró el Registro !"
   End If
End Sub


Private Sub Command18_Click()
 Frame1.Visible = False
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

  If Data1.Recordset.RecordCount = 0 Then Exit Sub
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
   
  
       
   Frame1.Visible = True
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
   If Not IsNull(Data1.Recordset("PRVCFAXACR")) Then
      Text9.text = Data1.Recordset("PRVCFAXACR")
   Else
      Text9.text = ""
   End If
      Text4 = IIf(IsNull(Data1.Recordset("PRVCLOCALI")), "", Data1.Recordset("PRVCLOCALI"))
      Text7 = IIf(IsNull(Data1.Recordset("PRVCPAISAC")), "", Data1.Recordset("PRVCPAISAC"))
      Text8 = IIf(IsNull(Data1.Recordset("PRVCGIROAC")), "", Data1.Recordset("PRVCGIROAC"))
      Text10 = IIf(IsNull(Data1.Recordset("PRVCREPRES")), "", Data1.Recordset("PRVCREPRES"))
      Text11 = IIf(IsNull(Data1.Recordset("PRVCCARREP")), "", Data1.Recordset("PRVCCARREP"))
      Text12 = IIf(IsNull(Data1.Recordset("PRVCTELREP")), "", Data1.Recordset("PRVCTELREP"))
      Text1.Enabled = False
      Text2.SetFocus
End Sub



Private Sub Command3_Click()
   
  Dim MENSAJE$
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
 
    
     
   MENSAJE$ = "Está seguro de Borrar ?" & Chr$(10)
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
 If Data1.Recordset.RecordCount = 0 Then Exit Sub

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
  Frame1.Visible = False
   Data1.RecordSource = "Select * from MAEPROV order by PRVCRUC"
   Data1.Refresh
End Sub
Private Sub Command7_Click()
   Unload Me
   
End Sub

Private Sub Command8_Click()
    If Text1.text = "" Then
      MsgBox "Ingrese el codigo", vbExclamation, "Proveedores"
      Text1.SetFocus
      Exit Sub
   End If
   Command1.Visible = True
   Command2.Visible = True
   Command3.Visible = True
   Command4.Visible = True
   Command5.Visible = True
   Command6.Visible = True
   Command7.Visible = True
   Command8.Visible = False
         
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
        Data1.Recordset("PRVCLOCALI") = Mid$(UCase$(Text4.text), 1, 8)
      Else
         Data1.Recordset("PRVCLOCALI") = " "
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
      Data1.Recordset("PRVCGIROAC") = IIf(Trim(Text8) <> "", Text8, " ")
       Data1.Recordset("PRVCPAISAC") = IIf(Trim(Text7) <> "", Text7, " ")
      Data1.Recordset("PRVCREPRES") = IIf(Trim(Text10) <> "", Text10, " ")
      Data1.Recordset("PRVCCARREP") = IIf(Trim(Text11) <> "", Text11, " ")
      Data1.Recordset("PRVCTELREP") = IIf(Trim(Text12) <> "", Text12, " ")
     Data1.Recordset("PRVDFECCRE") = Date
     
      
      If Text9.text <> "" Then
         Data1.Recordset("PRVCFAXACR") = Mid$(UCase$(Text9.text), 1, 15)
      Else
         Data1.Recordset("PRVCFAXACR") = " "
      End If
       
   Data1.Recordset.Update
   Data1.Refresh
   End If
   Text1 = ""
   Text2 = ""
   Text3 = ""
   Text4 = ""
   Text5 = ""
   Text6 = ""
   Text7 = ""
   Text8 = ""
   Text9 = ""
   Text10 = ""
   Tex11 = ""
   Text12 = ""
    Frame1.Visible = False
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
   Frame1.Visible = False
  
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
   
   Text1 = ""
   Text2 = ""
   Text3 = ""
   Text4 = ""
   Text5 = ""
   Text6 = ""
   Text7 = ""
   Text8 = ""
   Text9 = ""
   Text10 = ""
   Text11 = ""
   Text12 = ""
    Frame1.Visible = False
End Sub

Private Sub Form_Load()
Label1.Enabled = True
Label2.Enabled = True
Label3.Enabled = True
Label4.Enabled = True
Label5.Enabled = True
Label6.Enabled = True

Data1.DatabaseName = cRuta2
Init_ControlDBGrid DBGrid1
'Init_ControlDBGrid  (DBGrid1)
central FormMantProv


End Sub

Private Sub Text1_Change()
    
    If resp = "S" Then
    '  If Len(Text1.text) = 8 Then
'         criterio = "PRVCRUC = " & Chr$(34) + Text1.text + Chr$(34)
'         Data1.Recordset.FindFirst criterio
'         If Not Data1.Recordset.NoMatch Then
'            DBGrid1.Visible = False
'            MsgBox "El Tercero ya ha sido registrado !"
'            DBGrid1.Visible = False
'            Text1.text = ""
'            Text1.SetFocus
'         Else
'            Text1.Enabled = False
'
'            Text1.Visible = True
'
'            Text2.Visible = True
'            Text3.Visible = True
'            Text4.Visible = True
'            Text5.Visible = True
'            Text6.Visible = True
'            Text7.Visible = True
'
'            Text8.Visible = True
'            Text9.Visible = True
'            Text2.SetFocus
'
'         End If
'       End If
  End If
         
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
Dim criterio As String
If KeyAscii = 13 And Text1 <> "" Then
  criterio = "PRVCRUC = " & Chr$(34) + Text1.text + Chr$(34)
  Data1.Recordset.FindFirst criterio
  If Not Data1.Recordset.NoMatch Then
            DBGrid1.Visible = False
            MsgBox "El Proveedor ya ha sido registrado !"
            DBGrid1.Visible = False
            Text1.text = ""
            Text1.SetFocus
  Else
      siguientetab KeyAscii
 End If
End If
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
  siguientetab KeyAscii
End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
  siguientetab KeyAscii
End Sub

Private Sub Text12_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Command8.SetFocus
 End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
  siguientetab KeyAscii
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
   siguientetab KeyAscii
End Sub

Private Sub Text4_Change()
   'codigo postal
   If Len(Text4.text) = 6 Then
      criterio = "PRVCCODIGO = " & Chr$(34) + Text4.text + Chr$(34)
'      Data2.Recordset.FindFirst criterio
'      If Data2.Recordset.NoMatch Then
'         MsgBox "El Código Postal no ha sido registrado !"
'         Text4.text = ""
    
  End If
End Sub



Private Sub Text4_KeyPress(KeyAscii As Integer)
  siguientetab KeyAscii
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
   siguientetab KeyAscii
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
  siguientetab KeyAscii
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
  siguientetab KeyAscii
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
  siguientetab KeyAscii
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
  siguientetab KeyAscii
End Sub


Public Sub siguientetab(KeyAscii)
 If KeyAscii = 13 Then
   SendKeys "{tab}"
   KeyAscii = 0
End If
End Sub
