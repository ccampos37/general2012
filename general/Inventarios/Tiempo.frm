VERSION 5.00
Begin VB.Form Tiempo 
   Caption         =   "Tiempo Abierto"
   ClientHeight    =   855
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4725
   Icon            =   "Tiempo.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   855
   ScaleWidth      =   4725
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Tiempo de Uso"
      Height          =   375
      Left            =   3000
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "Tiempo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Command1_Click()
  Dim lngTickCount As Long
  lngTickCount = GetTickCount
  Call MsgBox("Has usado tu ordenador durante:" & vbCrLf & _
  " * " & CStr(lngTickCount) & " milisengundos, o " & vbCrLf & _
  " * " & CStr(lngTickCount / 1000) & " segundos, o " & vbCrLf & _
  " * " & CStr(lngTickCount / 60000) & " minutos", vbInformation)



End Sub

