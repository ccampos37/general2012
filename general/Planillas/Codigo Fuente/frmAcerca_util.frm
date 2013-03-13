VERSION 5.00
Begin VB.Form frmAcerca_util 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2640
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6330
   LinkTopic       =   "Form1"
   ScaleHeight     =   176
   ScaleMode       =   0  'User
   ScaleWidth      =   422
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   1425
      Top             =   2040
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00400000&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   5790
      ScaleHeight     =   240
      ScaleWidth      =   375
      TabIndex        =   3
      Top             =   90
      Width           =   375
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         ForeColor       =   &H8000000E&
         Height          =   165
         Left            =   120
         TabIndex        =   4
         Top             =   0
         Width           =   135
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   255
      Left            =   5805
      TabIndex        =   5
      Top             =   90
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   510
      Index           =   1
      Left            =   1065
      Picture         =   "frmAcerca_util.frx":0000
      Stretch         =   -1  'True
      Top             =   735
      Width           =   690
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Utilitario  Personalizado"
      ForeColor       =   &H80000018&
      Height          =   180
      Left            =   150
      TabIndex        =   2
      Top             =   180
      Width           =   1635
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00004040&
      BackStyle       =   1  'Opaque
      Height          =   435
      Left            =   0
      Top             =   -15
      Width           =   6345
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Este programa es exclusivo para los clientes de Enterprise Solutions S.A."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1125
      Index           =   0
      Left            =   2160
      TabIndex        =   1
      Top             =   450
      Width           =   3780
   End
   Begin VB.Image Image2 
      Height          =   510
      Index           =   0
      Left            =   1065
      Picture         =   "frmAcerca_util.frx":1E72
      Stretch         =   -1  'True
      Top             =   735
      Width           =   690
   End
   Begin VB.Image Image1 
      Height          =   1050
      Left            =   120
      Picture         =   "frmAcerca_util.frx":3CE4
      Stretch         =   -1  'True
      Top             =   690
      Width           =   705
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Utilitario  Personalizado de Reportes  Versión 1.0 "
      ForeColor       =   &H00800000&
      Height          =   180
      Index           =   0
      Left            =   2640
      TabIndex        =   0
      Top             =   2250
      Width           =   3435
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Este programa es exclusivo para los clientes de Enterprise Solutions S.A."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   1125
      Index           =   1
      Left            =   2175
      TabIndex        =   6
      Top             =   480
      Width           =   3780
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Utilitario  Personalizado de Reportes  Versión 1.0 "
      ForeColor       =   &H00808080&
      Height          =   180
      Index           =   1
      Left            =   2655
      TabIndex        =   7
      Top             =   2265
      Width           =   3435
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   255
      Top             =   1515
      Width           =   585
   End
   Begin VB.Image Image3 
      Height          =   390
      Left            =   975
      Picture         =   "frmAcerca_util.frx":418B
      Stretch         =   -1  'True
      Top             =   1095
      Width           =   660
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      Height          =   90
      Left            =   1215
      Shape           =   2  'Oval
      Top             =   1440
      Width           =   630
   End
End
Attribute VB_Name = "frmAcerca_util"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function CreateRectRgn& Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long)
Private Declare Function CreateRoundRectRgn& Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long)
Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetWindowRgn Lib "User32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long

Dim oldleft As Integer
Dim oldtop As Integer
Dim oldx As Integer
Dim oldy As Integer
Dim moving As Boolean

Private Sub Form_Load()
    Dim MeWidth As Long
    Dim MeHeight As Long
    MeWidth = Me.Width / Screen.TwipsPerPixelX
    MeHeight = Me.Height / Screen.TwipsPerPixelY
    lRet = CreateRoundRectRgn(0, 0, MeWidth, MeHeight, 20, 20)
    dl = SetWindowRgn(Me.hWnd, lRet, True)
    
    moving = False
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Y < 40 And Y > 3 Then
        moving = True
        oldleft = Me.Left
        oldtop = Me.TOP
        oldx = X * Screen.TwipsPerPixelX + Me.Left
        oldy = Y * Screen.TwipsPerPixelY + Me.TOP
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Picture1.BorderStyle = 0
   If moving Then
        thisx = X * Screen.TwipsPerPixelX + Me.Left
        thisy = Y * Screen.TwipsPerPixelY + Me.TOP
        Me.Left = oldleft + thisx - oldx
        Me.TOP = oldtop + thisy - oldy
    End If

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    moving = False
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Picture1.BorderStyle = 0
    If moving Then
        thisx = X * Screen.TwipsPerPixelX + Me.Left
        thisy = Y * Screen.TwipsPerPixelY + Me.TOP
        Me.Left = oldleft + thisx - oldx
        Me.TOP = oldtop + thisy - oldy
    End If
End Sub
Private Sub Label2_Click()
    Unload Me
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Picture1.BorderStyle = 1
End Sub

Private Sub Picture1_Click()
    Unload Me
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Picture1.BorderStyle = 1
End Sub

Private Sub Timer1_Timer()
    Static CONTADOR As Integer


    Image2(CONTADOR).Visible = False
        If CONTADOR = 1 Then
            CONTADOR = CONTADOR - 1
            'Label4(1).Visible = True
        Else
            'Label4(1).Visible = False
            CONTADOR = CONTADOR + 1
        End If
    Image2(CONTADOR).Visible = True
    


End Sub
