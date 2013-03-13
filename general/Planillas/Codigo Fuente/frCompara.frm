VERSION 5.00
Begin VB.Form frCompara 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Resultados de la Comparación"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5790
   Icon            =   "frCompara.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   5790
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&Cerrar"
      Height          =   375
      Left            =   4200
      TabIndex        =   3
      Top             =   4365
      Width           =   1455
   End
   Begin VB.ListBox List1 
      Height          =   3180
      Left            =   135
      TabIndex        =   2
      Top             =   1110
      Width           =   5535
   End
   Begin VB.Frame Frame1 
      Caption         =   "Base de Datos"
      Height          =   930
      Left            =   135
      TabIndex        =   0
      Top             =   135
      Width           =   5550
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         Height          =   300
         Left            =   135
         TabIndex        =   1
         Top             =   360
         Width           =   5265
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Programado por Daniel Yafac"
      Height          =   195
      Left            =   165
      TabIndex        =   4
      Top             =   4470
      Width           =   2085
   End
End
Attribute VB_Name = "frCompara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub
