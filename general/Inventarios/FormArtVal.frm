VERSION 5.00
Begin VB.Form FormArtVal 
   Caption         =   "Reporte  Valorizados"
   ClientHeight    =   4995
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6585
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   4995
   ScaleWidth      =   6585
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame5 
      Caption         =   "Tipo de Movimiento"
      Height          =   3615
      Left            =   360
      TabIndex        =   17
      Top             =   120
      Width           =   5655
      Begin VB.Frame FrameMon 
         Caption         =   "Moneda"
         Height          =   1335
         Left            =   3120
         TabIndex        =   28
         Top             =   600
         Visible         =   0   'False
         Width           =   2175
         Begin VB.OptionButton Option15 
            Caption         =   "Dolares"
            Height          =   255
            Left            =   240
            TabIndex        =   30
            Top             =   840
            Width           =   1335
         End
         Begin VB.OptionButton Option14 
            Caption         =   "Soles"
            Height          =   255
            Left            =   240
            TabIndex        =   29
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.OptionButton Option13 
         Caption         =   "Stock por Almacen"
         Height          =   375
         Left            =   360
         TabIndex        =   27
         Top             =   600
         Width           =   2415
      End
      Begin VB.Frame FrameMan 
         Height          =   1455
         Left            =   2160
         TabIndex        =   22
         Top             =   2040
         Visible         =   0   'False
         Width           =   3135
         Begin VB.TextBox Text6 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1920
            TabIndex        =   26
            Top             =   960
            Width           =   615
         End
         Begin VB.TextBox Text5 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1920
            TabIndex        =   24
            Top             =   480
            Width           =   615
         End
         Begin VB.Label Label8 
            Caption         =   "Al Almacen"
            Height          =   255
            Left            =   240
            TabIndex        =   25
            Top             =   960
            Width           =   1095
         End
         Begin VB.Label Label7 
            Caption         =   "Del Almacen"
            Height          =   375
            Left            =   240
            TabIndex        =   23
            Top             =   480
            Width           =   1575
         End
      End
      Begin VB.OptionButton Option12 
         Caption         =   "Familia"
         Height          =   375
         Left            =   360
         TabIndex        =   21
         Top             =   2520
         Width           =   1695
      End
      Begin VB.OptionButton Option11 
         Caption         =   "Linea"
         Height          =   375
         Left            =   360
         TabIndex        =   20
         Top             =   2040
         Width           =   1695
      End
      Begin VB.OptionButton Option10 
         Caption         =   "Grupo"
         Height          =   375
         Left            =   360
         TabIndex        =   19
         Top             =   1560
         Width           =   1695
      End
      Begin VB.OptionButton Option9 
         Caption         =   "Manualmente"
         Height          =   375
         Left            =   360
         TabIndex        =   18
         Top             =   1080
         Width           =   2175
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Marzo"
      Height          =   3615
      Left            =   360
      TabIndex        =   2
      Top             =   120
      Width           =   5655
      Begin VB.Frame Frame2 
         Caption         =   "Moneda"
         Height          =   1335
         Left            =   360
         TabIndex        =   8
         Top             =   2040
         Width           =   1935
         Begin VB.OptionButton Option1 
            Caption         =   "Soles"
            Height          =   375
            Left            =   360
            TabIndex        =   10
            Top             =   360
            Width           =   1215
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Dolares"
            Height          =   375
            Left            =   360
            TabIndex        =   9
            Top             =   720
            Width           =   1215
         End
      End
      Begin VB.TextBox Text3 
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
         Left            =   1080
         TabIndex        =   7
         Top             =   1560
         Width           =   1215
      End
      Begin VB.TextBox Text4 
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
         Left            =   3360
         TabIndex        =   6
         Top             =   1560
         Width           =   1335
      End
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
         Left            =   1080
         TabIndex        =   5
         Top             =   720
         Width           =   615
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4080
         TabIndex        =   4
         Text            =   "Combo1"
         Top             =   720
         Width           =   1335
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
         Left            =   3000
         TabIndex        =   3
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "De"
         Height          =   255
         Left            =   360
         TabIndex        =   16
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Al"
         Height          =   255
         Left            =   2280
         TabIndex        =   15
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Articulos"
         Height          =   375
         Left            =   360
         TabIndex        =   14
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "De"
         Height          =   255
         Left            =   360
         TabIndex        =   13
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label Label5 
         Caption         =   "Al"
         Height          =   255
         Left            =   2640
         TabIndex        =   12
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "Almacen"
         Height          =   255
         Left            =   360
         TabIndex        =   11
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.CommandButton Command7 
      Caption         =   "&Salir"
      Height          =   855
      Left            =   3360
      Picture         =   "FormArtVal.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3960
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   855
      Left            =   1920
      Picture         =   "FormArtVal.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3960
      Width           =   855
   End
End
Attribute VB_Name = "FormArtVal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
 If Frame5.Visible Then
     If FrameMan.Visible Then
        Unload Me
     End If
     Frame5.Visible = False
 Else
     Frame5.Visible = True
     Frame1.Visible = False
     'Codigo reporte
     Unload Me
 End If
 If Option10.Value Then
      Label3.Caption = " Por Grupos"
   ElseIf Option11.Value Then
   
      Label3.Caption = " Por Lineas"
   Else
      Label3.Caption = " Por Familias"
   End If
End Sub

Private Sub Command7_Click()
    If Frame5.Visible Then
        Unload Me
    Else
        Frame5.Visible = True
        Frame1.Visible = False
    End If
End Sub

Private Sub Form_Load()
  central FormArtVal
End Sub

Private Sub Option10_Click()
  Option10.Value = True
   FrameMan.Visible = False
   FrameMon.Visible = False
   
End Sub

Private Sub Option11_Click()
  Option11.Value = True
   FrameMan.Visible = False
   FrameMon.Visible = False
End Sub

Private Sub Option12_Click()
   Option12.Value = True
   FrameMan.Visible = False
   FrameMon.Visible = False
End Sub

Private Sub Option13_Click()
  Option13.Value = True
  FrameMon.Visible = True
End Sub

Private Sub Option9_Click()
  Option9.Value = True
  FrameMan.Visible = True
  FrameMon.Visible = False
  Text5.SetFocus
  
   
End Sub

Private Sub Text5_KeyUp(KeyCode As Integer, Shift As Integer)
 'Text6.SetFocus
 
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 And Text6.text <> "" Then
    Command1.SetFocus
 End If
End Sub

Private Sub Text6_KeyUp(KeyCode As Integer, Shift As Integer)
  ' Text5.SetFocus
End Sub
