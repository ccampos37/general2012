VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4980
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   ScaleHeight     =   4980
   ScaleWidth      =   6975
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   840
      Left            =   1380
      TabIndex        =   0
      Top             =   2385
      Width           =   4140
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vgcnx As ADODB.Connection



Private Sub Command1_Click()
    Dim ayuda As New dllayudaRecordSet.ClassFormAyuda
    Dim camp As Field, nreg As Long
    
    ayuda.SQLCadena = "select * from moneda"
    ayuda.PrimerCampo = "monedacodigo"
    Call ayuda.mostrar(vgcnx, camp, nreg)
End Sub

Private Sub Form_Load()
    Set vgcnx = New ADODB.Connection
    vgcnx.CursorLocation = adUseClient
    vgcnx.CommandTimeout = 0
    vgcnx.ConnectionTimeout = 0
    vgcnx.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=CONTAPRUEBA;Data Source=DESARROLLO4\SQL2000"
    vgcnx.Open
    'Call CtrAyu_Asiento(0).conexion(vgcnx)
    'Call CtrAyu_SubAsiento.conexion(vgcnx)
End Sub
