VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form Form10 
   Caption         =   "Form10"
   ClientHeight    =   5655
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6450
   LinkTopic       =   "Form10"
   ScaleHeight     =   5655
   ScaleWidth      =   6450
   StartUpPosition =   3  'Windows Default
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "Form10.frx":0000
      Height          =   3615
      Left            =   1080
      OleObjectBlob   =   "Form10.frx":0014
      TabIndex        =   0
      Top             =   960
      Width           =   4095
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5040
      Width           =   2580
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Db As Database, Rs As Recordset, Kar As Recordset
Private Sub Form_Load()
    Dim annomes As String
    Set Db = Workspaces(0).OpenDatabase(RUTA & NAMEBD)
    Set Rs = Db.OpenRecordset("lstKardex", dbOpenSnapshot)
    'set rs1 =db.OpenRecordset
    Dim xGrup As String, xSum As Double
    Db.Execute "DELETE * FROM  kardexaux"
    Set Kar = Db.OpenRecordset("kardexaux")
    xGrup = Rs!decodigo
    xSum = cantidadmes(xGrup, annomes)
    xAcu = 0
    annomes = "199910"
    Do While Not Rs.EOF
         If Left(Rs(8), 4) <> "0101" Then
            Rs.MoveNext
            'Loop
         Else
          If Rs!decodigo <> xGrup Then
            xGrup = Rs!decodigo
            xSum = cantidadmes(xGrup, annomes)
            xAcu = xSum
         End If
        xSum = xSum + Rs!ingresos - Rs!Salida
        Kar.AddNew
        Kar!c1 = Rs(0)
        Kar!c2 = Rs(1)
        Kar!c3 = Rs(2)
        Kar!c4 = Rs(3)
        Kar!c5 = Rs(4)
        Kar!c6 = Rs(5)
        Kar!c7 = Rs(6)
        Kar!c8 = xSum
        Kar!c9 = xAcu
        Kar.Update
        Rs.MoveNext
       End If
      Loop
    Set Data1.Recordset = Kar
    'MsgBox "hola"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Rs.Close
    Db.Close
End Sub


Function cantidadmes(codigo As String, annomes As String) As Double
 Dim Rsql As String
 Dim Rs As Recordset
 VGComp = "01"
 VGAlma = "01"
 Rsql = "select SMCANENT, SMCANSAL from MoResMes where SMCIA = '" & VGComp & "' AND SMALMA = '" & VGAlma & "'AND SMCODIGO= '" & codigo & "'AND SMMESPRO = '" & annomes & "'"  '
 Set Rs = Db.OpenRecordset(Rsql, dbOpenSnapshot)
 If Not Rs.EOF Then
   cantidadmes = Rs(0) - Rs(1)
 Else
   cantidadmes = 0
 End If
End Function
