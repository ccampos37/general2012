VERSION 5.00
Begin VB.Form frAcerca 
   BackColor       =   &H80000009&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Acerca de ..."
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4875
   Icon            =   "frAcerca.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   4875
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Créditos"
      Height          =   1320
      Left            =   232
      TabIndex        =   1
      Top             =   3240
      Width           =   4410
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fernando Cossio     -   Analista Programador"
         Height          =   195
         Left            =   135
         TabIndex        =   5
         Top             =   870
         Width           =   3090
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Equipo de Desarrollo  Planillas"
         ForeColor       =   &H00008000&
         Height          =   195
         Left            =   900
         TabIndex        =   3
         Top             =   225
         Width           =   2130
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fernando Cossio     -   Jefe de Proyecto"
         Height          =   195
         Left            =   150
         TabIndex        =   2
         Top             =   585
         Width           =   2790
      End
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&Cerrar"
      Default         =   -1  'True
      Height          =   315
      Left            =   180
      TabIndex        =   0
      Top             =   2565
      Width           =   1200
   End
   Begin VB.Label Label9 
      BackColor       =   &H80000012&
      Caption         =   "  MS SQL Server 2000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   2940
      TabIndex        =   4
      Top             =   30
      Width           =   1935
   End
End
Attribute VB_Name = "frAcerca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
End Sub
