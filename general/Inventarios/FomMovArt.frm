VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FormMovArt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Movimiento Anual de Articulo"
   ClientHeight    =   3675
   ClientLeft      =   945
   ClientTop       =   1950
   ClientWidth     =   5010
   Icon            =   "FomMovArt.frx":0000
   LinkTopic       =   "Form12"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   5010
   Begin VB.CommandButton Command7 
      Caption         =   "&Salir"
      Height          =   735
      Left            =   2595
      Picture         =   "FomMovArt.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2775
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   735
      Left            =   1395
      Picture         =   "FomMovArt.frx":0D0C
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2775
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Height          =   2505
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4575
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "FomMovArt.frx":114E
         Left            =   1290
         List            =   "FomMovArt.frx":1150
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   360
         Width           =   3060
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1290
         TabIndex        =   1
         Top             =   1350
         Width           =   2175
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   300
         Left            =   1290
         TabIndex        =   9
         Top             =   855
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy"
         Format          =   53477379
         CurrentDate     =   36437
         MaxDate         =   401768
         MinDate         =   36161
      End
      Begin VB.Label Label9 
         Caption         =   "Año"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   825
         Width           =   735
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1290
         TabIndex        =   8
         Top             =   1830
         Width           =   3015
      End
      Begin VB.Label Label1 
         Caption         =   "Descripción"
         Height          =   255
         Left            =   210
         TabIndex        =   7
         Top             =   1815
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Almacén"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label8 
         Caption         =   "Articulo"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   210
         TabIndex        =   3
         Top             =   1335
         Width           =   735
      End
   End
End
Attribute VB_Name = "FormMovArt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit
Public Fecha As String
'Dim db As Database
Dim almacen As String
Dim almacenAnt As String
Dim rs As ADODB.Recordset

Private Sub Combo1_Click()
 rs.MoveFirst
 rs.Move Combo1.ListIndex
 almacen = Format(rs(0), "00")
End Sub

Private Sub Command1_Click()
      If Existe_cod_art(Text1) = "" Then
           Exit Sub
      End If
     If Text1 = "" Then
       MsgBox "Ingrese el Código del artículo", vbInformation, "Inventarios"
       Text1.SetFocus
       Exit Sub
    End If
    Fecha = Year(DTPicker1)
    frmMovAnual.Frame1.Caption = Mid(Text1 + " " + Label3, 1, 37)
    frmMovAnual.Show 1
End Sub

Private Sub Carga_Almacen()
   Dim rsql As String
   Dim I As Integer
   rsql = "Select TAALMA,TADESCRI FROM TabAlm "
   Set rs = New ADODB.Recordset
  rs.Open rsql, VGcnx, adOpenStatic
  Do While Not rs.EOF
     Combo1.AddItem (rs(1))
     rs.MoveNext
     If rs.EOF Then Exit Do
  Loop
   rs.MoveFirst
    For I = 0 To rs.RecordCount - 1
      If rs(0) = VGAlma Then
        Combo1.ListIndex = I
        Exit For
      Else
        rs.MoveNext
      End If
    Next
'   If rS.EOF Then
'     Combo1.AddItem ("Todos")
'   End If
  
  
End Sub

Private Sub Command7_Click()
 Unload Me
End Sub

Private Sub Form_Load()
   DTPicker1 = Date
   Carga_Almacen
   central FormMovArt
   VGForm1 = 11
End Sub

Private Sub Text1_DblClick()
 Label3 = ""
 VGForm1 = 11
 Text1 = ""
 almacenAnt = VGAlma
 VGAlma = almacen
 FormAyuArt1.Show 1
   If Text1 <> "" Then
        Command1.SetFocus
   End If
VGAlma = almacenAnt
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 46 Then Label3 = ""
   If KeyCode = 112 Then Text1_DblClick
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 And Text1 <> "" Then
      If Existe_cod_art(Text1) <> "" Then
             Command1.SetFocus
      End If
   End If
End Sub

Function Existe_cod_art(text As TextBox) As String
 
 Dim rs As New ADODB.Recordset
 Dim rsql As String
  rsql = "select  ACODIGO,Adescri FROM maeart where ACODIGO = '" & text & "'" '
  'Set db = Workspaces(0).OpenDatabase(cRuta2)
  'Set RS = db.OpenRecordset(RSQL, dbOpenSnapshot)
  
  Set rs = VGcnx.Execute(rsql)
  If Not rs.EOF Then
    Existe_cod_art = rs(0)
    Label3 = Mid(rs(1), 1, 20)
  Else
    MsgBox "El tipo de codigo no existe !", vbOKOnly + vbExclamation, "Error"
    Existe_cod_art = ""
  End If
   rs.Close
End Function

