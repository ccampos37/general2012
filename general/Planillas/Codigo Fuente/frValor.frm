VERSION 5.00
Object = "{C5FF36B7-A67E-11D3-9D9D-E6F193F7F854}#9.0#0"; "APLICTXT.OCX"
Begin VB.Form frValor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Valor"
   ClientHeight    =   1305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2805
   Icon            =   "frValor.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1305
   ScaleWidth      =   2805
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   1530
      TabIndex        =   3
      Top             =   780
      Width           =   1095
   End
   Begin VB.CommandButton cmAceptar 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   345
      Left            =   195
      TabIndex        =   2
      Top             =   780
      Width           =   1095
   End
   Begin AplisetControlText.Aplitext xValor 
      Height          =   300
      Left            =   945
      TabIndex        =   1
      Top             =   255
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   529
      MaxLength       =   8
      Text            =   "0"
      TipoDato        =   "N"
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Valor"
      Height          =   195
      Left            =   345
      TabIndex        =   0
      Top             =   300
      Width           =   360
   End
End
Attribute VB_Name = "frValor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMACEPTAR_CLICK()
    VPTAREA = xValor.Text
    Unload Me
End Sub

Private Sub CMCANCELAR_CLICK()
        VPTAREA = "0"
    Unload Me
End Sub

