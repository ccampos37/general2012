VERSION 5.00
Begin VB.Form frmRepListadoDiferenciasCompras 
   Caption         =   "Listar Diferencias Compras"
   ClientHeight    =   3480
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5580
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3480
   ScaleWidth      =   5580
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H00800000&
      Height          =   1665
      Left            =   45
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "frmRepListadoDiferenciasCompras.frx":0000
      Top             =   90
      Width           =   5490
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   405
      Left            =   2805
      TabIndex        =   1
      Top             =   2445
      Width           =   1500
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   405
      Left            =   1275
      TabIndex        =   0
      Top             =   2445
      Width           =   1500
   End
End
Attribute VB_Name = "frmRepListadoDiferenciasCompras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdaceptar_Click()
  Call imprimir
End Sub

Private Sub CmdCancelar_Click()
   Unload Me
End Sub

Sub imprimir()
Dim arrform(0) As Variant, arrparm(5) As Variant
On Error GoTo Imprime
    Screen.MousePointer = 11
    arrparm(0) = VGParamSistem.BDEmpresaCT
    arrparm(1) = VGParamSistem.BDEmpresa
    arrparm(2) = VGParamSistem.Servidor
    arrparm(3) = VGParamSistem.Anoproceso
    arrparm(4) = Format(VGParamSistem.Mesproceso, "00")
    Call ImpresionRptProc("rptListarDiferenciasCompras.rpt", arrform, arrparm, , "Listar Diferencias")
    Screen.MousePointer = 1
    Exit Sub
Imprime:
Screen.MousePointer = 1
MsgBox Err.Description

End Sub

