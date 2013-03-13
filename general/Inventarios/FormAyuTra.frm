VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form FormAyuTransa 
   Caption         =   "Ayuda de Transaccion"
   ClientHeight    =   4785
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6345
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4785
   ScaleWidth      =   6345
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command8 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   735
      Left            =   3600
      Picture         =   "FormAyuTra.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3960
      Width           =   840
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   735
      Left            =   1920
      Picture         =   "FormAyuTra.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3960
      Width           =   840
   End
   Begin VB.Frame Frame1 
      Caption         =   "Buscar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   3735
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   6015
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
         Left            =   1680
         TabIndex        =   0
         Top             =   480
         Width           =   2175
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "FormAyuTra.frx":0884
         Height          =   2655
         Left            =   120
         OleObjectBlob   =   "FormAyuTra.frx":0898
         TabIndex        =   5
         Top             =   960
         Width           =   5775
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   240
         Picture         =   "FormAyuTra.frx":1287
         Top             =   360
         Width           =   480
      End
      Begin VB.Label Label1 
         Caption         =   "Codigo"
         Height          =   255
         Left            =   960
         TabIndex        =   4
         Top             =   480
         Width           =   615
      End
   End
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
      RecordSource    =   "TabTransa"
      Top             =   4080
      Visible         =   0   'False
      Width           =   1140
   End
End
Attribute VB_Name = "FormAyuTransa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Combo1_Click()
  
  End Sub

Private Sub Command1_Click()
   ' Data1.Refresh
   If VGForm = 20 Then
    If frmkardexDoc.Text1 = "" Then
          frmkardexDoc.Text1 = Data1.Recordset.Fields(1)
          frmkardexDoc.lbltrans1 = Data1.Recordset.Fields("tt_descri")
    Else
           frmkardexDoc.Text2 = Data1.Recordset.Fields(1)
           frmkardexDoc.lbltrans2 = Data1.Recordset.Fields("tt_descri")
    End If
 Else
    If VGForm = 6 Then
      FrmGuiaSal.TxTransa = Data1.Recordset.Fields(1)
     Else
      FormRegistro.TxTransa = Data1.Recordset.Fields(1)
    End If
  End If
  Unload Me
End Sub

Private Sub Command8_Click()
  Unload Me
  
End Sub

Private Sub DBGrid1_Click()
  'Command1_Click
End Sub

Private Sub DBGrid1_DblClick()
  Command1_Click
End Sub

Private Sub DBGrid1_SelChange(Cancel As Integer)
  'Command1_Click
End Sub

Private Sub Form_Load()
  Data1.DatabaseName = cRuta2
 If VGForm = 6 Then VGRegEnt = 2
 If VGRegEnt = 1 Then
   Data1.RecordSource = "SELECT * FROM TabTransa WHERE TT_TIPMOV = 'I' ORDER BY TT_CODMOV"
 Else
   Data1.RecordSource = "SELECT * FROM TabTransa WHERE TT_TIPMOV = 'S' ORDER BY TT_CODMOV"
 End If
 Init_ControlDBGrid DBGrid1
'  Combo1.ListIndex = 0
 ' Label1.Caption = Combo1.Text
 ' central FormAyuTransa
   AlinearAyuda Me
End Sub

Private Sub Text1_Change()
 Dim ncar As String
  ncar = Str$(Len(Text1))
  criterio = "MID$(TT_CODMOV,1," + ncar + ") = '" & Text1 & "'"
  Data1.Recordset.FindFirst criterio
 
End Sub



