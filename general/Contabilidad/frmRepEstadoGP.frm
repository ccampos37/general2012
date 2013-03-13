VERSION 5.00
Begin VB.Form xxxfrmRepEstadoGP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Estado de Ganancias y Pérdidas"
   ClientHeight    =   2205
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5040
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2205
   ScaleWidth      =   5040
   Begin VB.CommandButton cmdBotones 
      Caption         =   "&Aceptar"
      Height          =   345
      Index           =   0
      Left            =   1312
      TabIndex        =   3
      Top             =   1605
      Width           =   1050
   End
   Begin VB.CommandButton cmdBotones 
      Caption         =   "&Cancelar"
      Height          =   345
      Index           =   1
      Left            =   2662
      TabIndex        =   2
      Top             =   1605
      Width           =   1050
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo de Reporte"
      Height          =   900
      Left            =   0
      TabIndex        =   0
      Top             =   315
      Width           =   5025
      Begin VB.ComboBox cboTipoReporte 
         Height          =   315
         Left            =   45
         TabIndex        =   1
         Top             =   300
         Width           =   4935
      End
   End
End
Attribute VB_Name = "xxxfrmRepEstadoGP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
  Call LlenarcboTipoReporte
  Width = 5160
  Height = 2610
End Sub

Private Sub cboTipoReporte_Click()
  cmdBotones(0).SetFocus
  Select Case cboTipoReporte.ListIndex
    Case 0:
    
    Case 1
    
    Case 2:
  
  End Select
End Sub

Sub LlenarcboTipoReporte()
  cboTipoReporte.Clear
  cboTipoReporte.AddItem "Estado de Ganancias y Pérdidas"
  cboTipoReporte.AddItem "Sustento de Estado de Ganancias y Pérdidas"
  cboTipoReporte.AddItem "Resumen de Estado de Ganacias y Pérdidas"
End Sub

Private Sub cmdBotones_Click(Index As Integer)
  Select Case Index
    Case 0
    
    Case 1: Unload Me
  
  End Select

End Sub
