VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frWait 
   BorderStyle     =   0  'None
   ClientHeight    =   1395
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2790
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1395
   ScaleWidth      =   2790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar Progreso 
      Height          =   135
      Left            =   510
      TabIndex        =   2
      Top             =   870
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   1
      Min             =   1
      Max             =   5
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   80
      Left            =   105
      Top             =   330
   End
   Begin VB.Shape Shape2 
      Height          =   1170
      Left            =   0
      Top             =   225
      Width           =   2775
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Procesando Tareas ..."
      Height          =   195
      Left            =   600
      TabIndex        =   1
      Top             =   630
      Width           =   1575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Espere ..."
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   45
      TabIndex        =   0
      Top             =   30
      Width           =   675
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000002&
      BackStyle       =   1  'Opaque
      Height          =   240
      Left            =   0
      Top             =   0
      Width           =   2775
   End
End
Attribute VB_Name = "frWait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
    Progreso.Value = 1
    Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
    Progreso.Value = Progreso.Value + 1
    If Progreso.Value >= 5 Then
        Unload Me
    End If
End Sub
