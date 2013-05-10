VERSION 5.00
Begin VB.Form FrmgenerasaldosAnaliticos 
   Caption         =   "Generacion de Movimientos de Cta.Cte"
   ClientHeight    =   1920
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6360
   LinkTopic       =   "Form3"
   ScaleHeight     =   1920
   ScaleWidth      =   6360
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      Height          =   495
      Left            =   4440
      TabIndex        =   5
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   4440
      TabIndex        =   4
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   495
      Left            =   2040
      TabIndex        =   3
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   495
      Left            =   2040
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Año Actual"
      Height          =   495
      Left            =   720
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Año Anterior"
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "FrmgenerasaldosAnaliticos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
 Dim SQL As String
  On Error GoTo xx
    Screen.MousePointer = 11
    VGCNx.BeginTrans
    
    Set VGCommandoSP = New ADODB.Command
    VGCommandoSP.ActiveConnection = VGGeneral
    VGCommandoSP.CommandType = adCmdStoredProc
    VGCommandoSP.CommandText = "ct_GeneraCtaCteApertura_pro"
    VGCommandoSP.Parameters.Refresh
    With VGCommandoSP
        .Parameters("@base") = VGParamSistem.BDEmpresa
        .Parameters("@empresa") = VGParametros.empresacodigo
        .Parameters("@annoact") = VGParamSistem.Anoproceso
        .Parameters("@annopas") = Trim$(VGParamSistem.Anoproceso - 1)
        .Parameters("@NombrePC") = VGcomputer
        .Execute
    End With
    VGCNx.CommitTrans
    Screen.MousePointer = 1
    MsgBox "Se Genero la Cuenta Corriente de Apertura del Año " & VGParamSistem.Anoproceso, vbInformation
    Command1.Enabled = False
    Exit Sub
xx:
    Screen.MousePointer = 1
    VGCNx.RollbackTrans
    MsgBox "No se pudo Aperturar la Cuenta Corriente " & Chr(13) & err.Description, vbExclamation
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Text1.Text = VGParamSistem.Anoproceso - 1
Text2.Text = VGParamSistem.Anoproceso
End Sub
