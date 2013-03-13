VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmPrcSaldos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Proceso de Regeneracion de Saldos"
   ClientHeight    =   3045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6075
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   6075
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   645
      Left            =   1980
      Picture         =   "FrmPrcSaldos.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2130
      Width           =   795
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancelar"
      Height          =   645
      Left            =   3255
      Picture         =   "FrmPrcSaldos.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2130
      Width           =   795
   End
   Begin VB.Frame Frame1 
      Height          =   1875
      Left            =   210
      TabIndex        =   0
      Top             =   150
      Width           =   5655
      Begin VB.Frame Frame2 
         Height          =   645
         Left            =   150
         TabIndex        =   5
         Top             =   1110
         Width           =   5355
         Begin MSComctlLib.ProgressBar Bar1 
            Height          =   315
            Left            =   90
            TabIndex        =   6
            Top             =   210
            Width           =   5145
            _ExtentX        =   9075
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   1
            Appearance      =   0
         End
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   270
         Width           =   4335
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   4110
         TabIndex        =   7
         Top             =   810
         Width           =   1365
      End
      Begin VB.Label Label1 
         Caption         =   "Almacen"
         Height          =   255
         Index           =   0
         Left            =   270
         TabIndex        =   4
         Top             =   360
         Width           =   735
      End
   End
End
Attribute VB_Name = "FrmPrcSaldos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset



Private Sub Combo1_Change()
Command1.Visible = True
End Sub

Private Sub Command1_Click()

  Dim Text2 As String
  
  Text2 = "" & Combo1.text
VGCNx.BeginTrans
   Set VGCommandoSP = New ADODB.Command
   VGCommandoSP.ActiveConnection = VGgeneral
   VGCommandoSP.CommandType = adCmdStoredProc
   VGCommandoSP.CommandText = "al_RegeneraSaldos_pro"
   VGCommandoSP.Parameters.Refresh
   With VGCommandoSP
       .Parameters("@base") = VGParamSistem.BDEmpresa
       .Parameters("@alma") = Left(Text2, 2)
       .Execute
   End With
   VGCNx.CommitTrans
  MsgBox "Proceso Terminado Satisfactoriamente..!!", vbInformation, "AVISO"
  Command1.Visible = False
  Exit Sub
nerror:
    If Err Then
        If nflag = 1 Then
            VGCNx.RollbackTrans
        End If
        MsgBox "Error : " & Err.Number & "-" & Err.Description
        Err = 0
        Exit Sub
        
    End If
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
  Dim rsc As New ADODB.Recordset
  
  Combo1.Clear
  Set rsc = VGCNx.Execute("select TAALMA,TADESCRI from tabalm where empresacodigo='" & VGparametros.empresacodigo & "' and puntovtacodigo='" & VGparametros.puntovta & "'")
  If rsc.RecordCount > 0 Then
      rsc.MoveFirst
      Do Until rsc.EOF
        Combo1.AddItem rsc.Fields(0) & "-" & IIf(IsNull(rsc.Fields(1)), "****", rsc.Fields(1))
        rsc.MoveNext
      Loop
  End If
  rsc.Close
  Set rsc = Nothing
  
End Sub
