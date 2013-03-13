VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form FormMantDoc 
   Caption         =   "Mantenimiento de Documento"
   ClientHeight    =   4170
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5895
   LinkTopic       =   "Form11"
   ScaleHeight     =   4170
   ScaleWidth      =   5895
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   480
      TabIndex        =   11
      Top             =   2880
      Width           =   4815
      Begin VB.CommandButton Command19 
         Caption         =   "&Salir"
         Height          =   675
         Left            =   3000
         Picture         =   "FormMantDoc.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   240
         Width           =   775
      End
      Begin VB.CommandButton Command21 
         Caption         =   "Grabar"
         Height          =   675
         Left            =   720
         Picture         =   "FormMantDoc.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   240
         Width           =   775
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2775
      Left            =   480
      TabIndex        =   6
      Top             =   120
      Width           =   4815
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1920
         TabIndex        =   10
         Text            =   "Text2"
         Top             =   1440
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1920
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Descripcion"
         Height          =   255
         Left            =   600
         TabIndex        =   8
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Codigo"
         Height          =   255
         Left            =   600
         TabIndex        =   7
         Top             =   960
         Width           =   1335
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Adicionar"
      Height          =   675
      Left            =   600
      Picture         =   "FormMantDoc.frx":058C
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3120
      Width           =   775
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Modificar"
      Height          =   675
      Left            =   1800
      Picture         =   "FormMantDoc.frx":0896
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3120
      Width           =   775
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Eliminar"
      Height          =   675
      Left            =   2880
      Picture         =   "FormMantDoc.frx":0CD8
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3120
      Width           =   775
   End
   Begin VB.CommandButton Command7 
      Caption         =   "&Salir"
      Height          =   675
      Left            =   6480
      Picture         =   "FormMantDoc.frx":111A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   0
      Width           =   775
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Inventarios.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   615
      Left            =   4800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "TabTipDoc"
      Top             =   3840
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton Command8 
      Caption         =   "&Salir"
      Height          =   675
      Left            =   4080
      Picture         =   "FormMantDoc.frx":1264
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3120
      Width           =   775
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "FormMantDoc.frx":13AE
      Height          =   2775
      Left            =   480
      OleObjectBlob   =   "FormMantDoc.frx":13C2
      TabIndex        =   1
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "FormMantDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 Dim modifica As Boolean

Private Sub Command19_Click()
  Frame1.Visible = False
End Sub

Private Sub Command2_Click()
 modifica = True
 limpia
 Frame1.Visible = True
End Sub

Private Sub Command21_Click()
If Not modifica Then
 Data1.Recordset("tdo_tipdoc") = Text1
End If
Data1.Recordset("tdo_tipdoc") = Text2
Data1.Refresh
End Sub

Private Sub Command4_Click()
  limpia
  Frame1.Visible = True
  modifica = False
  
End Sub

Private Sub Command8_Click()
 Unload Me
End Sub

Private Sub Form_Load()
 Data1.DatabaseName = cRuta2
 limpia
 'direccionar la ruta mdb
End Sub
Private Sub limpia()
  Text1 = ""
  Text2 = ""
  
End Sub

