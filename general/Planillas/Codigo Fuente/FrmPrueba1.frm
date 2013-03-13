VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Migrar"
      Height          =   315
      Left            =   60
      TabIndex        =   0
      Top             =   2745
      Width           =   1065
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsPrueba As New ADODB.Recordset
Dim RsConcep As New ADODB.Recordset
Dim i As Integer
Private Sub Command1_Click()
    dt1.CN1.Open
    RsPrueba.Open "Prueba", dt1.CN1, adOpenDynamic, adLockReadOnly
    RsConcep.Open "Conceptos", dt1.CN1, adOpenDynamic, adLockOptimistic
    For i = 0 To RsPrueba.Fields.Count - 1
        RsConcep.AddNew
        RsConcep.Fields(0).Value = RsPrueba.Fields(i).Name
        RsConcep.Update
    Next
End Sub
