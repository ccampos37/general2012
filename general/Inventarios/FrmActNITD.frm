VERSION 5.00
Begin VB.Form FrmActNITD 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Actualizar Costo de Notas de Ingreso por Transferencia"
   ClientHeight    =   1875
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5625
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1875
   ScaleWidth      =   5625
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdExit 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   495
      Left            =   3240
      TabIndex        =   1
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton CmdActCostoDestino 
      Caption         =   "Actualizar"
      Default         =   -1  'True
      Height          =   495
      Left            =   1200
      TabIndex        =   0
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"FrmActNITD.frx":0000
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   5655
   End
End
Attribute VB_Name = "FrmActNITD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
