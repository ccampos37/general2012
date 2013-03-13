VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form FrmDetImp 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Detalles de la Importacion"
   ClientHeight    =   2955
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   5895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   315
      Left            =   2490
      TabIndex        =   1
      Top             =   2640
      Width           =   1035
   End
   Begin RichTextLib.RichTextBox Rich 
      Height          =   2565
      Left            =   60
      TabIndex        =   0
      Top             =   45
      Width           =   5730
      _ExtentX        =   10107
      _ExtentY        =   4524
      _Version        =   393217
      ReadOnly        =   -1  'True
      TextRTF         =   $"FrmDetImp.frx":0000
   End
End
Attribute VB_Name = "FrmDetImp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdSalir_Click()
    Unload Me
End Sub
