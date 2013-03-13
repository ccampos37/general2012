VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4500
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4500
   ScaleWidth      =   6000
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox ListaTipos 
      Height          =   2010
      Left            =   240
      TabIndex        =   6
      Top             =   2040
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tamaño"
      Height          =   2295
      Left            =   3120
      TabIndex        =   0
      Top             =   1920
      Width           =   2175
      Begin VB.OptionButton Tamaño24 
         Caption         =   "24"
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   1440
         Width           =   855
      End
      Begin VB.OptionButton Tamaño16 
         Caption         =   "16"
         Height          =   495
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   1335
      End
      Begin VB.OptionButton Tamaño12 
         Caption         =   "12"
         Height          =   495
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Tipos"
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Etiqueta 
      Alignment       =   2  'Center
      Caption         =   "Texto de prueba"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   600
      TabIndex        =   4
      Top             =   240
      Width           =   4335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim N
For N = 0 To Screen.FontCount - 1
ListaTipos.AddItem Screen.Fonts(N)
Next N
End Sub

Private Sub ListaTipos_Click()
Etiqueta.Font.Name = ListaTipos.Text
End Sub

Private Sub Tamaño12_Click()
Etiqueta.Font.Size = Tamaño12.Caption
End Sub

Private Sub Tamaño16_Click()
Etiqueta.Font.Size = Tamaño16.Caption
End Sub


Private Sub Tamaño24_Click()
Etiqueta.Font.Size = Tamaño24.Caption
End Sub
