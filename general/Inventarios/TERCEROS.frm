VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Catálogo de Clientes"
   ClientHeight    =   5895
   ClientLeft      =   240
   ClientTop       =   1005
   ClientWidth     =   9030
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleMode       =   0  'User
   ScaleWidth      =   9119.851
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   240
      TabIndex        =   28
      Top             =   3600
      Width           =   8415
      Begin VB.CommandButton Command19 
         Caption         =   "&Salir"
         Height          =   675
         Left            =   4800
         Picture         =   "TERCEROS.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   360
         Width           =   775
      End
      Begin VB.CommandButton Command21 
         Caption         =   "Aceptar"
         Height          =   675
         Left            =   3240
         Picture         =   "TERCEROS.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   360
         Width           =   775
      End
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Inventarios.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   7560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "MAECLI"
      Top             =   360
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton Command20 
      Caption         =   "Direcciones"
      Height          =   675
      Left            =   4080
      Picture         =   "TERCEROS.frx":058C
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   4920
      Visible         =   0   'False
      Width           =   767
   End
   Begin VB.CommandButton Command18 
      Caption         =   "&Cancelar Buscar"
      Height          =   675
      Left            =   6480
      Picture         =   "TERCEROS.frx":09CE
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   4560
      Visible         =   0   'False
      Width           =   767
   End
   Begin VB.CommandButton Command17 
      Caption         =   "Por &R. Social"
      Height          =   675
      Left            =   4080
      MaskColor       =   &H00C0FFFF&
      Picture         =   "TERCEROS.frx":0B18
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   4920
      Visible         =   0   'False
      Width           =   767
   End
   Begin VB.CommandButton Command16 
      Caption         =   "Por &RUC"
      Height          =   675
      Left            =   1680
      MaskColor       =   &H00C0FFFF&
      Picture         =   "TERCEROS.frx":0F5A
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   4320
      Visible         =   0   'False
      Width           =   767
   End
   Begin VB.CommandButton Command15 
      Caption         =   "&Cancelar Orden"
      Height          =   675
      Left            =   6480
      Picture         =   "TERCEROS.frx":139C
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   4920
      Visible         =   0   'False
      Width           =   767
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Por &Razón Social"
      Height          =   675
      Left            =   4080
      Picture         =   "TERCEROS.frx":14E6
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   4920
      Visible         =   0   'False
      Width           =   767
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Por &Código"
      Height          =   675
      Left            =   1680
      Picture         =   "TERCEROS.frx":1630
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   4920
      Visible         =   0   'False
      Width           =   767
   End
   Begin VB.CommandButton Command6 
      Caption         =   "&Ordenar "
      Height          =   675
      Left            =   6480
      Picture         =   "TERCEROS.frx":177A
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   4920
      Width           =   767
   End
   Begin VB.CommandButton Command12 
      Caption         =   "C&ancelar Imprimir"
      Height          =   675
      Left            =   6480
      Picture         =   "TERCEROS.frx":1A84
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   4920
      Visible         =   0   'False
      Width           =   767
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Por &Razón Social"
      Height          =   675
      Left            =   4080
      Picture         =   "TERCEROS.frx":1BCE
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   4920
      Visible         =   0   'False
      Width           =   767
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Por &Código"
      Height          =   675
      Left            =   1680
      Picture         =   "TERCEROS.frx":1ED8
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   4920
      Visible         =   0   'False
      Width           =   767
   End
   Begin VB.CommandButton Command9 
      Caption         =   "&Cancelar"
      Height          =   675
      Left            =   6480
      Picture         =   "TERCEROS.frx":21E2
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4920
      Visible         =   0   'False
      Width           =   767
   End
   Begin VB.CommandButton Command8 
      Caption         =   "&Grabar"
      Height          =   675
      Left            =   1680
      Picture         =   "TERCEROS.frx":232C
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4920
      Visible         =   0   'False
      Width           =   767
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "TERCEROS.frx":2636
      Height          =   4335
      Left            =   240
      OleObjectBlob   =   "TERCEROS.frx":264A
      TabIndex        =   16
      Top             =   120
      Width           =   8415
   End
   Begin VB.CommandButton Command7 
      Caption         =   "&Salir"
      Height          =   675
      Left            =   7680
      Picture         =   "TERCEROS.frx":3529
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4920
      Width           =   767
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&Imprimir"
      Height          =   675
      Left            =   5280
      Picture         =   "TERCEROS.frx":3673
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4920
      Width           =   767
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Buscar"
      Height          =   675
      Left            =   4080
      Picture         =   "TERCEROS.frx":3AB5
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4920
      Width           =   767
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Eliminar"
      Height          =   675
      Left            =   2880
      Picture         =   "TERCEROS.frx":3EF7
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4920
      Width           =   767
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Modificar"
      Height          =   675
      Left            =   1680
      Picture         =   "TERCEROS.frx":4339
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4920
      Width           =   767
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Adicionar"
      Height          =   675
      Left            =   480
      Picture         =   "TERCEROS.frx":477B
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4920
      Width           =   767
   End
   Begin VB.Frame Frame2 
      Height          =   4215
      Left            =   360
      TabIndex        =   31
      Top             =   240
      Width           =   8295
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   2040
         TabIndex        =   7
         Top             =   3720
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2040
         TabIndex        =   2
         Top             =   840
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2040
         TabIndex        =   3
         Top             =   1440
         Visible         =   0   'False
         Width           =   4455
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   2040
         TabIndex        =   4
         Top             =   2040
         Visible         =   0   'False
         Width           =   4455
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   2040
         TabIndex        =   5
         Top             =   2640
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   2040
         TabIndex        =   6
         Top             =   3240
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   2040
         TabIndex        =   1
         Top             =   360
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   5280
         TabIndex        =   8
         Top             =   3240
         Visible         =   0   'False
         Width           =   1575
      End
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   6840
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   262150
      End
      Begin VB.Label Label7 
         Caption         =   "Vendedor:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   480
         TabIndex        =   40
         Top             =   3720
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "R.U.C:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   480
         TabIndex        =   39
         Top             =   840
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Razon Social"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   480
         TabIndex        =   38
         Top             =   1440
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Dirección"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   480
         TabIndex        =   37
         Top             =   2040
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Zona Postal:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   480
         TabIndex        =   36
         Top             =   2640
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "Teléfono"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   480
         TabIndex        =   35
         Top             =   3240
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "Codigo"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   480
         TabIndex        =   34
         Top             =   360
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "Fax:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4200
         TabIndex        =   33
         Top             =   3240
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label10 
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4200
         TabIndex        =   32
         Top             =   2520
         Visible         =   0   'False
         Width           =   2295
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim resp As String
Dim buscar As Boolean
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
   'Label6.Visible = True
   Label7.Visible = True
   Label8.Visible = True
   Label9.Visible = True
   
  
   Text1.Visible = True
   Text2.Visible = True
   Text3.Visible = True
   Text4.Visible = True
   Text5.Visible = True
   Text7.Visible = True
   Text8.Visible = True
   Text9.Visible = True
      
Text1.SetFocus
End Sub



Private Sub Command10_Click()
'   Data1.RecordSource = "SELECT * FROM MAECLI ORDER BY CNUMRUC"
'   Data1.Refresh
   Dim CODIGO2 As String
   CODIGO2 = "ENTERPRISSE S.A."
   CrystalReport1.ReportFileName = RUTA & "reporteinv\catclie.rpt"
   Ubi_Tab CrystalReport1
   CrystalReport1.DiscardSavedData = True
   CrystalReport1.Formulas(0) = "EMPRESA='" & CODIGO2 & "'"
   CrystalReport1.Destination = crptToWindow
   CrystalReport1.Action = 1
End Sub

Private Sub Command11_Click()
'   Data1.RecordSource = "SELECT * FROM MAECLI ORDER BY CNOMCLI"
'   Data1.Refresh
   Dim CODIGO2 As String
   CODIGO2 = "ENTERPRISSE S.A."
   CrystalReport1.ReportFileName = RUTA & "reporteinv\catclied.rpt"
    Ubi_Tab CrystalReport1
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

src = InputBox$("Ingrese RUC", "Búsqueda")
ncar = Str$(Len(src))
criterio = "MID$(CNUMRUC,1," + ncar + ") = " & Chr$(34) + src + Chr$(34)
Data1.Recordset.FindFirst criterio
If Data1.Recordset.NoMatch Then
   MsgBox "No se encontró el Registro !"
End If



End Sub

Private Sub Command17_Click()
DBGrid1.Enabled = True
DBGrid1.Visible = True

src = InputBox$("Ingrese Nombre del Tercero", "Búsqueda")
   ncar = Str$(Len(src))
   criterio = "MID$(CNOMCLI,1," + ncar + ") = " & Chr$(34) + src + Chr$(34)
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
   
   
   
  End Sub

Private Sub Command19_Click()
If buscar Then
   Command17_Click
Else
   'Form3.Show
    Unload Me
End If
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
   
   'Command20.Visible = True
    
  
   Label1.Visible = True
   Label2.Visible = True
   Label3.Visible = True
   Label4.Visible = True
   Label5.Visible = True
   'Label6.Visible = True
   Label7.Visible = True
   Label8.Visible = True
   Label9.Visible = True
   
   
   Text8 = Data1.Recordset("CCODCLI")
   If Not IsNull(Data1.Recordset("CNUMRUC")) Then
       Text1.text = Data1.Recordset("CNUMRUC")
   End If
   If Not IsNull(Data1.Recordset("CNOMCLI")) Then
      Text2.text = Data1.Recordset("CNOMCLI")
   Else
      Text2.text = ""
    End If
   
   If Not IsNull(Data1.Recordset("CDIRCLI")) Then
      Text3.text = Data1.Recordset("CDIRCLI")
   Else
      Text3.text = ""
   End If
   
   If Not IsNull(Data1.Recordset("CTELEFO")) Then
      Text5.text = Data1.Recordset("CTELEFO")        '   verificar el adicionar
    Else
      Text5.text = ""
   End If
   If Not IsNull(Data1.Recordset("CZONPOS")) Then
      Text4.text = Data1.Recordset("CZONPOS")
   Else
      Text4.text = ""
   End If
   
  If Not IsNull(Data1.Recordset("CVENDE")) Then
      Text7.text = Data1.Recordset("CVENDE")
   Else
      Text7.text = ""
   End If

'
'    If Not IsNull(Data1.Recordset(7)) Then
'      Text7.text = Data1.Recordset(7)
'   Else
'     Text7.text = ""
'  End If
'
'
   If Not IsNull(Data1.Recordset("CNUMFAX")) Then
      Text9.text = Data1.Recordset("CNUMFAX")
   Else
      Text9.text = ""
   End If
   
   
'   If Not IsNull(Data1.Recordset(5)) Then
'     Text9.text = Data1.Recordset(5)
'   Else
'     Text9.text = ""
'    End If
   
   
   Text1.Visible = True
   Text1.Enabled = False
   Text2.Visible = True
   Text3.Visible = True
   Text4.Visible = True
   Text5.Visible = True
   
   Text7.Visible = True
   Text8.Visible = True
   Text9.Visible = True
   
   
   
  
   Text2.SetFocus
End Sub

Private Sub Command20_Click()
 '  Form4.Show
End Sub

Private Sub Command21_Click()
If VGForm = 6 Then
   FormGuiaSal.Text5 = Data1.Recordset.Fields("CNUMRUC")
   FormGuiaSal.Text6 = Data1.Recordset.Fields("CNOMCLI")
   FormGuiaSal.Text7 = Data1.Recordset.Fields("CDIRCLI")
Else
   FormRegistro.Text7 = Data1.Recordset.Fields("CNUMRUC")
   FormRegistro.lblClie = Data1.Recordset.Fields("CNOMCLI")
End If
Unload Me
  
End Sub

Private Sub Command22_Click()
  If buscar Then
    buscar = False
    Command21.Caption = "&Aceptar"
    Command19.Caption = "&Salir"
    Command22.Caption = "&Buscar"
  Else
    buscar = False
    Command21.Caption = "Por &RUC"
    Command19.Caption = "Por Razon &Social"
    Command22.Caption = "Cancelar &Buscar"
  End If
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
 
    
     
   MENSAJE$ = "Está seguro de Borrar el Tercero?" & Chr$(10)
   MENSAJE$ = MENSAJE$ & Data1.Recordset("CNOMCLI")
   If MsgBox(MENSAJE$, 33, "Borra el Proveedor ?") <> 1 Then Exit Sub
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
   Data1.RecordSource = "Select * from MAECLI order by CNUMRUC"
   Data1.Refresh
End Sub
Private Sub Command7_Click()
   Unload Form2
   
   
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
   

   
      
   If Text1 <> "" Then
      If resp = "S" Then
         Data1.Recordset.AddNew
         Data1.Recordset("CCODCLI") = Trim(Text8.text)
         Data1.Recordset("CNUMRUC") = UCase$(Text1.text)
      Else
         Data1.Recordset.Edit
      End If
      
      
      If Text2 <> "" Then
         Data1.Recordset("CNOMCLI") = Mid$(UCase$(Text2.text), 1, 50)
      Else
         Data1.Recordset("CNOMCLI") = " "
      End If
      
      If Text3 <> "" Then
         Data1.Recordset("CDIRCLI") = Mid$(UCase$(Text3.text), 1, 50)
      Else
         Data1.Recordset("CDIRCLI") = " "
      End If
      
      If Text4 <> "" Then
        Data1.Recordset("CZONPOS") = Mid$(UCase$(Text4.text), 1, 5)
      Else
         Data1.Recordset("CZONPOS") = " "
      End If
      
      
      If Text5 <> "" Then
         Data1.Recordset("CTELEFO") = Mid$(UCase$(Text5.text), 1, 8)
      Else
         Data1.Recordset("CTELEFO") = " "
      End If
     
      
        If Text7 <> "" Then
         Data1.Recordset("CVENDE") = Mid$(UCase$(Text7.text), 1, 15)
         Else
         Data1.Recordset("CVENDE") = " "
         
      End If
      
      
'        If Text7 <> "" Then
'         Data1.Recordset(7) = Text7
'      Else
'         Data1.Recordset(7) = " "
'      End If
      
      
        If Text9 <> "" Then
         Data1.Recordset("CNUMFAX") = Mid$(UCase$(Text9.text), 1, 50)
      Else
         Data1.Recordset("CNUMFAX") = " "
      End If
      
      
'        If Text9 <> "" Then
'         Data1.Recordset(5) = Mid$(UCase$(Text9.text), 1, 8)
'      Else
'         Data1.Recordset(5) = " "
'      End If
      
      
    End If
            
   
   Data1.Recordset.Update
   Data1.Refresh
   
   
   Label1.Visible = False
   Label2.Visible = False
   Label3.Visible = False
   Label4.Visible = False
   Label5.Visible = False
   ' Label6.Visible = False
   Label7.Visible = False
   Label8.Visible = False
   Label9.Visible = False
   
   
   Text1.Enabled = True
   
   Text1.Visible = False
   Text2.Visible = False
   Text3.Visible = False
   Text4.Visible = False
   Text5.Visible = False
   Text7.Visible = False
   Text8.Visible = False
   Text9.Visible = False
   
   
   
   Text1 = ""
   Text2 = ""
   Text3 = ""
   Text4 = ""
   Text5 = ""
  
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
   Data1.RecordSource = "Select * from MAECLI order by CNOMCLI"
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
   'Label6.Visible = False
   Label7.Visible = False
   Label8.Visible = False
   Label9.Visible = False
   
   Text1.Enabled = True
   Text1.Visible = False
   Text2.Visible = False
   Text3.Visible = False
   Text4.Visible = False
   Text5.Visible = False
   Text7.Visible = False
   Text8.Visible = False
   Text9.Visible = False
   
   Text1 = ""
   Text2 = ""
   Text3 = ""
   Text4 = ""
   Text5 = ""
   Text7 = ""
   Text8 = ""
   Text9 = ""
End Sub



Private Sub Data2_Validate(Action As Integer, Save As Integer)

End Sub

Private Sub Form_Load()
Label1.Enabled = True
Label2.Enabled = True
Label3.Enabled = True
Label4.Enabled = True
Label5.Enabled = True
'Label6.Enabled = True
Label7.Enabled = True

Label1.Visible = False
Label2.Visible = False
Label3.Visible = False
Label4.Visible = False
Label5.Visible = False
'Label6.Visible = False
Data1.DatabaseName = RUTA & NAMEBD
Init_ControlDBGrid DBGrid1
central Form2
'*********** FALTA RECORSET  *******
If VGAyuClie Then
   'Frame1.Visible = True
   Command1.Visible = False
   Command2.Visible = False
   Command3.Visible = False
   Command4.Visible = True
   Command5.Visible = False
   Command6.Visible = False
   Command7.Visible = False
   Command21.Visible = True
   Command19.Visible = True
Else
     Frame1.Visible = False
End If

End Sub

Private Sub Text1_Change()
    
    If resp = "S" Then
      If Len(Text1.text) = 8 Then
         criterio = "CNUMRUC = " & Chr$(34) + Text1.text + Chr$(34)
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
            Text7.Visible = True
            Text8.Visible = True
            Text9.Visible = True
            Text2.SetFocus
         
         End If
       End If
   End If
         
End Sub

Private Sub Text4_Change()
   
'   If Len(Text4.text) = 6 Then
'      CRITERIO = "POS_CODIGO = " & Chr$(34) + Text4.text + Chr$(34)
'      Data2.Recordset.FindFirst CRITERIO
'      If Data2.Recordset.NoMatch Then
'         MsgBox "El Código Postal no ha sido registrado !"
'         Text4.text = ""
'      Else
'         Text4.text = UCase$(Text4.text)
'         Label10.Caption = Data2.Recordset(1)
'         Label10.Visible = True
'      End If
'   End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
 ' siguientetab
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
  ' siguientetab
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
 ' siguientetab
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
' siguientetab
End Sub


