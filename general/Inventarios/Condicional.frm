VERSION 5.00
Begin VB.Form Condicional 
   Caption         =   "Form2"
   ClientHeight    =   6030
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4485
   LinkTopic       =   "Form2"
   Picture         =   "Condicional.frx":0000
   ScaleHeight     =   6030
   ScaleWidth      =   4485
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Condicional"
      Height          =   495
      Left            =   2880
      TabIndex        =   0
      Top             =   2040
      Width           =   1455
   End
End
Attribute VB_Name = "Condicional"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim Operacion
Operacion = Introduce(Valor)
End Sub

Private Sub Command2_Click()
End
End Sub

