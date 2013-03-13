VERSION 5.00
Begin VB.Form FrmOpPlanG 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Imprimir Planilla General"
   ClientHeight    =   1635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4665
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1635
   ScaleWidth      =   4665
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancelar"
      Height          =   390
      Left            =   3420
      TabIndex        =   5
      Top             =   885
      Width           =   1140
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   390
      Left            =   3420
      TabIndex        =   4
      Top             =   345
      Width           =   1155
   End
   Begin VB.Frame Frame1 
      Caption         =   "Opciones del Reporte"
      Height          =   1350
      Left            =   90
      TabIndex        =   0
      Top             =   120
      Width           =   3165
      Begin VB.OptionButton Option1 
         Caption         =   "Todos los Trabajadores"
         Height          =   240
         Index           =   2
         Left            =   75
         TabIndex        =   3
         Top             =   975
         Width           =   2880
      End
      Begin VB.OptionButton Option1 
         Caption         =   "No declarados al PDT"
         Height          =   240
         Index           =   1
         Left            =   75
         TabIndex        =   2
         Top             =   690
         Width           =   2265
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Declarados al PDT Sunat"
         Height          =   240
         Index           =   0
         Left            =   75
         TabIndex        =   1
         Top             =   420
         Width           =   2490
      End
   End
End
Attribute VB_Name = "FrmOpPlanG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Aceptar As Boolean
Public op As Integer

Private Sub Command1_Click()
    Aceptar = True
    Unload Me
End Sub

Private Sub Command2_Click()
    Aceptar = False
    Unload Me
End Sub

Private Sub Form_Load()
    Aceptar = False
    Option1(0).Value = True
End Sub

Private Sub Option1_Click(Index As Integer)
    op = Index
End Sub
