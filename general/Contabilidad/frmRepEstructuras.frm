VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRepEstructuras 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Estructuras"
   ClientHeight    =   2250
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5070
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   5070
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   195
      Left            =   825
      TabIndex        =   4
      Top             =   1935
      Width           =   3420
      _ExtentX        =   6033
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Frame Frame1 
      Caption         =   "Seleccionar Tipo"
      Height          =   840
      Left            =   0
      TabIndex        =   2
      Top             =   120
      Width           =   5070
      Begin VB.ComboBox cboTipoReporte 
         Height          =   315
         Left            =   45
         TabIndex        =   3
         Top             =   435
         Width           =   4980
      End
   End
   Begin VB.CommandButton cmdBotones 
      Caption         =   "&Aceptar"
      Height          =   345
      Index           =   0
      Left            =   1335
      TabIndex        =   1
      Top             =   1365
      Width           =   1050
   End
   Begin VB.CommandButton cmdBotones 
      Caption         =   "&Cancelar"
      Height          =   345
      Index           =   1
      Left            =   2685
      TabIndex        =   0
      Top             =   1365
      Width           =   1050
   End
End
Attribute VB_Name = "frmRepEstructuras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboTipoReporte_Click()
  cmdBotones(0).SetFocus
End Sub

Private Sub Form_Load()
  Call LlenarcboTipoReporte
  ProgressBar1.Visible = False
  Call ConfiguraForm
End Sub

Sub LlenarcboTipoReporte()
  cboTipoReporte.Clear
  cboTipoReporte.AddItem "Balance General"
  cboTipoReporte.AddItem "Estado de Ganáncias y Pérdidas"
End Sub

Sub ConfiguraForm()
  Width = 5190
  Height = 2655
  'Left = (MDIPrincipal.Width - Me.Width) / 2
  'Top = ((MDIPrincipal.Height - Me.Height) / 2) - 600
End Sub

Private Sub cmdBotones_Click(Index As Integer)
  Select Case Index
    Case 0:
      Call prueba
    
    Case 1: Unload Me
  
  End Select

End Sub

Sub prueba()
 ProgressBar1.Min = 1
 ProgressBar1.Max = 100
 ProgressBar1.Visible = True
 
 ProgressBar1.Value = 10
 
 ProgressBar1.Value = 30
 
 ProgressBar1.Value = 50
 
 ProgressBar1.Value = 60

 'ProgressBar1.Value = 100

 'ProgressBar1.Visible = False

End Sub
