VERSION 5.00
Begin VB.Form IngDato 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "  "
   ClientHeight    =   1095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3690
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1095
   ScaleWidth      =   3690
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "& Aceptar"
      Height          =   315
      Left            =   1935
      TabIndex        =   1
      Top             =   675
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   345
      Left            =   1125
      TabIndex        =   0
      Top             =   135
      Width           =   2445
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   285
      Left            =   105
      TabIndex        =   2
      Top             =   150
      Width           =   810
   End
End
Attribute VB_Name = "IngDato"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub COMMAND1_CLICK()
If Len(Trim(Text1.Text)) > 0 Then
    Select Case LlamaFrm
    Case 1
        ModPlan.DatoTrabajador.CodigoTrab = Text1.Text
    Case 2
        ModPlan.DatoTrabajador.CTABANCO = Text1.Text
    Case 3
        ModPlan.DatoTrabajador.CtaCte = Text1.Text
    Case 4
        ModPlan.DatoTrabajador.CUSPP = Text1.Text
    Case 5
        ModPlan.DatoTrabajador.Departamento = Text1.Text
    End Select
    Unload Me
Else
    MsgBox "Es necesario que ingrese el dato faltante", vbCritical, "Informacion"
    Text1.SetFocus
End If
End Sub

Private Sub Text1_GotFocus()
Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)
End Sub

Private Sub Text1_KeyPress(KEYASCII As Integer)
If KEYASCII = 13 Then
    SendKeys "{Tab}"
    KEYASCII = 0
End If
End Sub
