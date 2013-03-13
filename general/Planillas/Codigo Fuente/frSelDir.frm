VERSION 5.00
Begin VB.Form frSelDir 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Seleccionar carpeta destino"
   ClientHeight    =   4515
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3750
   Icon            =   "frSelDir.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   3750
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmCancela 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   1988
      TabIndex        =   5
      Top             =   4065
      Width           =   1230
   End
   Begin VB.CommandButton cmAcepta 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   345
      Left            =   533
      TabIndex        =   4
      Top             =   4065
      Width           =   1230
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   390
      Width           =   3510
   End
   Begin VB.DirListBox Dir1 
      Height          =   2340
      Left            =   135
      TabIndex        =   0
      Top             =   1005
      Width           =   3510
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   3
      Left            =   135
      TabIndex        =   7
      Top             =   3660
      Width           =   3495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Carpeta Seleccionada"
      Height          =   195
      Index           =   2
      Left            =   135
      TabIndex        =   6
      Top             =   3435
      Width           =   1575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Carpetas disponibles"
      Height          =   195
      Index           =   1
      Left            =   135
      TabIndex        =   3
      Top             =   780
      Width           =   1455
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Unidades en el Sistema"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   150
      Width           =   1665
   End
End
Attribute VB_Name = "frSelDir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmAcepta_Click()
    vpTarea = Dir1.Path
    Unload Me
End Sub

Private Sub cmCancela_Click()
    Unload Me
End Sub

Private Sub Dir1_Change()
    Label1(3).Caption = Dir1.Path
End Sub

Private Sub Drive1_Change()
    On Error GoTo ErrDrive
    Dir1.Path = Drive1.Drive
    Exit Sub
ErrDrive:
    MsgBox Err.Description
    Exit Sub
End Sub

Private Sub Form_Activate()
    vpTarea = ""
    Dir1_Change
End Sub
