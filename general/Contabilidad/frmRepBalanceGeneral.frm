VERSION 5.00
Begin VB.Form frmRepBalanceGeneral 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Balance General"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5550
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   5550
   Begin VB.Frame Frame1 
      Caption         =   "Tipo de Reporte"
      Height          =   915
      Left            =   0
      TabIndex        =   6
      Top             =   195
      Width           =   5550
      Begin VB.ComboBox cboTipoReporte 
         Height          =   315
         Left            =   60
         TabIndex        =   8
         Top             =   330
         Width           =   5430
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Criterios de Filtro"
      Height          =   1035
      Left            =   0
      TabIndex        =   2
      Top             =   1215
      Width           =   5550
      Begin VB.TextBox txtLinea 
         Height          =   285
         Left            =   3720
         TabIndex        =   7
         Top             =   585
         Width           =   1755
      End
      Begin VB.ComboBox cboNiveles 
         Height          =   315
         Left            =   1665
         TabIndex        =   4
         Top             =   195
         Width           =   3810
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "Todas las líneas"
         Height          =   240
         Left            =   45
         TabIndex        =   3
         Top             =   615
         Width           =   1800
      End
      Begin VB.Label Label2 
         Caption         =   "Línea"
         Height          =   240
         Left            =   2985
         TabIndex        =   9
         Top             =   630
         Width           =   465
      End
      Begin VB.Label Label1 
         Caption         =   "Nivel Cuenta"
         Height          =   270
         Left            =   90
         TabIndex        =   5
         Top             =   285
         Width           =   1140
      End
   End
   Begin VB.CommandButton cmdBotones 
      Caption         =   "&Aceptar"
      Height          =   345
      Index           =   0
      Left            =   1575
      TabIndex        =   1
      Top             =   2700
      Width           =   1050
   End
   Begin VB.CommandButton cmdBotones 
      Caption         =   "&Cancelar"
      Height          =   345
      Index           =   1
      Left            =   2925
      TabIndex        =   0
      Top             =   2700
      Width           =   1050
   End
End
Attribute VB_Name = "frmRepBalanceGeneral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
  Call LlenarcboTipoReporte
  Call ConfiguraForm
End Sub

Sub ConfiguraForm()
  Me.Width = 5670
  Me.Height = 3660
  'Left = (MDIPrincipal.Width - Me.Width) / 2
  'Top = ((MDIPrincipal.Height - Me.Height) / 2) - 600
End Sub

Private Sub cboTipoReporte_Click()
  Select Case cboTipoReporte.ListIndex
    Case 0:
    
    Case 1
    
    Case 2:
  
  End Select
End Sub

Sub LlenarcboTipoReporte()
  cboTipoReporte.Clear
  cboTipoReporte.AddItem "Balance General"
  cboTipoReporte.AddItem "Sustento del Balance General"
  cboTipoReporte.AddItem "Resumen de Balance General"
End Sub

Private Sub cmdBotones_Click(Index As Integer)
  Select Case Index
    Case 0
    
    Case 1: Unload Me
  
  End Select

End Sub

Sub Llenarcboniveles()
 Dim I As Integer
 For I = 1 To VGnumniveles
   cboNiveles.AddItem "NIVEL " & Format(I, "0#")
 Next
End Sub
