VERSION 5.00
Begin VB.Form Seleccion 
   Caption         =   "Selección"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "Seleccion.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Seleccion"
      Height          =   1935
      Left            =   1080
      TabIndex        =   0
      Top             =   480
      Width           =   2415
      Begin VB.OptionButton Option2 
         Caption         =   "Escoge 2"
         Height          =   615
         Left            =   480
         TabIndex        =   2
         Top             =   1080
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Escoge 1"
         Height          =   495
         Left            =   480
         TabIndex        =   1
         Top             =   480
         Width           =   1455
      End
   End
End
Attribute VB_Name = "Seleccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Option1_Click()
If Option1.Enabled Then
MsgBox ("Eres bueno")
End If
End Sub

Private Sub Option2_Click()
If Option2.Enabled Then
MsgBox ("Eres Malo")
End If
End Sub
