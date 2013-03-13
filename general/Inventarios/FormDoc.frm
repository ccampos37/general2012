VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form FormDoc 
   Caption         =   "Documentos"
   ClientHeight    =   8025
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9345
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8025
   ScaleWidth      =   9345
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command8 
      Caption         =   "&Salir"
      Height          =   675
      Left            =   2880
      Picture         =   "FormDoc.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3240
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
      Left            =   3915
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "TabTipDoc"
      Top             =   2670
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   675
      Left            =   1440
      Picture         =   "FormDoc.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3240
      Width           =   775
   End
   Begin VB.Frame Frame1 
      Height          =   3135
      Left            =   270
      TabIndex        =   2
      Top             =   75
      Width           =   4935
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "FormDoc.frx":0884
         Height          =   2775
         Left            =   120
         OleObjectBlob   =   "FormDoc.frx":0898
         TabIndex        =   3
         Top             =   240
         Width           =   4695
      End
   End
End
Attribute VB_Name = "FormDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If VGForm = 5 Then
 'MsgBox "VGFORM" & VGForm
 FormRegistro.Text6.text = Data1.Recordset.Fields(0)
Else
 FrmGuiaSal.Text3.text = Data1.Recordset.Fields(0)
End If
 
  Unload Me
End Sub

Private Sub Command8_Click()
  Unload Me
End Sub

Private Sub Form_Load()
   Data1.DatabaseName = cRuta2
   AlinearAyuda Me
   Init_ControlDBGrid DBGrid1
End Sub
