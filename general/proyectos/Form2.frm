VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form2 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   1890
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3510
   LinkTopic       =   "Form2"
   ScaleHeight     =   1890
   ScaleWidth      =   3510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3240
      Top             =   1650
   End
   Begin ComctlLib.ImageList ImgL 
      Left            =   2370
      Top             =   1680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   1
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form2.frx":0000
            Key             =   "Logo"
         EndProperty
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   2520
      Stretch         =   -1  'True
      Top             =   480
      Width           =   945
   End
   Begin VB.Label LblEmp 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   240
      Left            =   120
      TabIndex        =   3
      Top             =   60
      Width           =   60
   End
   Begin VB.Label LblCar 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Administrador"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   240
      Left            =   90
      TabIndex        =   2
      Top             =   900
      Width           =   1335
   End
   Begin VB.Label LblUsu 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Elliot Gonzales"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   240
      Left            =   90
      TabIndex        =   1
      Top             =   540
      Width           =   1425
   End
   Begin VB.Label LblSaludo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E88651&
      Height          =   240
      Left            =   90
      TabIndex        =   0
      Top             =   1260
      Width           =   60
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E88651&
      BackStyle       =   1  'Opaque
      Height          =   405
      Left            =   -30
      Top             =   -30
      Width           =   3585
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim c As Integer
Private Sub Form_Load()
c = 1
If Right(Time, 4) = "a.m." Then
    LblSaludo.Caption = "Buenos Dias"
ElseIf Right(Time, 4) = "p.m." And CInt(Left(Time, 2)) <= 6 Then
    LblSaludo.Caption = "Buenos Tardes"
Else
    LblSaludo.Caption = "Buenos Noches"
End If

LblEmp.Caption = Trim(VGParametros.nomempresa)
Image1.Picture = ImgL.ListImages.Item(1).Picture

End Sub

Private Sub Timer1_Timer()
c = c + 1
If c = 8 Then
    SlideForm Form2, 1, 200
End If
End Sub
