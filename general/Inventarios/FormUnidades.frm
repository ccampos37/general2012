VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Unidades"
   ClientHeight    =   4044
   ClientLeft      =   3600
   ClientTop       =   1368
   ClientWidth     =   4248
   Icon            =   "FormUnidades.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4044
   ScaleWidth      =   4248
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   2865
      Width           =   3975
      Begin VB.CommandButton Command1 
         Caption         =   "&Aceptar"
         Height          =   675
         Left            =   720
         Picture         =   "FormUnidades.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   775
      End
      Begin VB.CommandButton Command8 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         Height          =   675
         Left            =   2520
         Picture         =   "FormUnidades.frx":0D0C
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   775
      End
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "FormUnidades.frx":114E
      Height          =   2775
      Left            =   150
      OleObjectBlob   =   "FormUnidades.frx":1162
      TabIndex        =   0
      Top             =   60
      Width           =   3975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Select Case VGForm1
   Case 1            'desde   1 al  case 3 estuvo con comentario
     FormCreacion.Text2.text = Data2.Recordset.Fields(1)
   Case 2 'varform = "FormCreacionSin"
     FormCreacionSin.Text3.text = Data2.Recordset.Fields(1)
   Case 3 'articulos
      FormArticulos.lblUnidad.Caption = Data2.Recordset.Fields(1)
   Case 4  'UNIDADES
      FrmArUniMed.Text4 = Data2.Recordset.Fields("UM_ABREV")  ' VGform 4
      
  End Select
  '
  VGabrev = Data2.Recordset.Fields("UM_ABREV")
  FrmArUniMed.Label7 = Data2.Recordset.Fields("UM_NOMBRE")
  Unload Me
End Sub

Private Sub Command8_Click()
   Unload Me
End Sub

Private Sub Form_Activate()
DBGrid1.SetFocus
End Sub

Private Sub Form_Load()
  'central Form1
AlinearAyuda Me
Init_ControlDBGrid DBGrid1
Data2.DatabaseName = cRuta2
Data2.RecordSource = "Select * from TABUNIMED  order by UM_ABREV"
Data2.Refresh
Command1.Default = True

DBGrid1.Refresh
End Sub
