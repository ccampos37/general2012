VERSION 5.00
Begin VB.Form FormSalImp 
   Caption         =   "Reporte"
   ClientHeight    =   5010
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6855
   LinkTopic       =   "Form1"
   ScaleHeight     =   5010
   ScaleWidth      =   6855
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FrameRep 
      Caption         =   "Rep"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   720
      TabIndex        =   2
      Top             =   360
      Width           =   5415
      Begin VB.Frame Frame1 
         Height          =   1335
         Left            =   360
         TabIndex        =   3
         Top             =   360
         Width           =   4215
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1440
            TabIndex        =   5
            Top             =   360
            Width           =   1455
         End
         Begin VB.TextBox Text2 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1440
            TabIndex        =   4
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label Label2 
            Caption         =   "Inico "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            TabIndex        =   7
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label3 
            Caption         =   "Fin"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            TabIndex        =   6
            Top             =   840
            Width           =   855
         End
      End
   End
   Begin VB.CommandButton Command7 
      Caption         =   "&Salir"
      Height          =   855
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3960
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   855
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3960
      Width           =   855
   End
End
Attribute VB_Name = "FormSalImp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   Unload Me
   
   
End Sub

Private Sub Command7_Click()
  Unload Me
  
End Sub

Private Sub Form_Load()
  Dim val As Integer
    val = 1
   If val = 1 Then
      FrameRep.Caption = " Por Articulos"
   Else
      FrameRep.Caption = " Por Familias"
   End If
  central FormSalImp
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      FormAyuArt.Show
   End If
   
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
      FormAyuArt.Show
   End If
End Sub
