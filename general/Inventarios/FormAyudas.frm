VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form FormAyudas 
   Caption         =   "Form11"
   ClientHeight    =   4965
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7650
   LinkTopic       =   "Form11"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   7650
   StartUpPosition =   3  'Windows Default
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Inventarios.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   615
      Left            =   5160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "TabAYUD"
      Top             =   3960
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Frame Frame1 
      Caption         =   "Buscar"
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
      Height          =   1095
      Left            =   360
      TabIndex        =   2
      Top             =   0
      Width           =   7080
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2040
         TabIndex        =   3
         Top             =   480
         Width           =   3060
      End
      Begin VB.Label Label1 
         Caption         =   "Codigo"
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   855
      Left            =   1830
      Picture         =   "FormAyudas.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3930
      Width           =   855
   End
   Begin VB.CommandButton Command8 
      Caption         =   "&Salir"
      Height          =   855
      Left            =   4815
      Picture         =   "FormAyudas.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3915
      Width           =   855
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "FormAyudas.frx":058C
      Height          =   2415
      Left            =   360
      OleObjectBlob   =   "FormAyudas.frx":05A0
      TabIndex        =   5
      Top             =   1320
      Width           =   7095
   End
End
Attribute VB_Name = "FormAyudas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  FrmGuiaSal.Text10.text = Data1.Recordset.Fields(1)
  Unload Me
End Sub

Private Sub Command8_Click()
  Unload Me
  
End Sub

Private Sub DBGrid1_Click()
    'Command1_Click
End Sub

Private Sub Form_Load()
   'VGEmp = "01"
  
   codayu = "22"
   FormAyudas.Caption = "Forma de Pagos "
  Data1.DatabaseName = cRuta2
  Data1.RecordSource = "SELECT * FROM TABAYU where TCOD= '" & codayu & "'   "
  Data1.Refresh
 Init_ControlDBGrid DBGrid1
'  Combo1.ListIndex = 0
 ' Label1.Caption = Combo1.Text
  AlinearAyuda Me
   
End Sub

Private Sub Text1_Change()
 Dim ncar As String
  ncar = Str$(Len(Text1))
  criterio = "Left(TCLAVE," & ncar & ") = '" & Text1 & "'"
  Data1.Recordset.FindFirst criterio
  
'   If Data1.Recordset.NoMatch Then
'      MsgBox "No se encontró el Registro !"
'   End If
End Sub

