VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Wait 
   BackColor       =   &H8000000A&
   ClientHeight    =   1560
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   4665
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   1560
   ScaleWidth      =   4665
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   3240
      Top             =   1800
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   720
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   360
      TabIndex        =   2
      Top             =   1080
      Width           =   45
   End
   Begin VB.Label Label1 
      Caption         =   "Espere un momento por favor ..."
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   240
      Width           =   3375
   End
End
Attribute VB_Name = "Wait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
 ProgressBar1.Min = 0
End Sub

Public Sub CargarWait(Mx As Integer, tipo As Integer)
 Select Case tipo
  Case 1:
    ProgressBar1.Scrolling = ccScrollingSmooth
  Case Else
    ProgressBar1.Scrolling = ccScrollingStandard
 End Select
 ProgressBar1.Max = Mx + 1
 Me.Show 1
 Wait.Refresh
End Sub

Public Sub Inc()
 If ProgressBar1.Value = ProgressBar1.Max Then
  Unload Me
 Else
  ProgressBar1.Value = ProgressBar1.Value + 1
 End If
End Sub

Public Sub PonLabel(S As String)
 Label2.Caption = S
 Label2.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
 Label2.Caption = "Proceso Finalizado"
End Sub

Public Sub Mx(m As Integer)
 ProgressBar1.Max = m
End Sub

