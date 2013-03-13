VERSION 5.00
Begin VB.Form FrmImpresora 
   Caption         =   "Seleccionar Impresora"
   ClientHeight    =   2130
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5100
   ForeColor       =   &H8000000F&
   Icon            =   "FrmImpresora.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2130
   ScaleWidth      =   5100
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2085
      Left            =   45
      TabIndex        =   7
      Top             =   0
      Width           =   5010
      Begin VB.ComboBox CboTama 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   330
         Left            =   1080
         TabIndex        =   4
         Top             =   1665
         Width           =   2760
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   330
         Left            =   3950
         TabIndex        =   6
         Top             =   585
         Width           =   960
      End
      Begin VB.CommandButton CmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   330
         Left            =   3950
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   180
         Width           =   960
      End
      Begin VB.ComboBox CboOrie 
         Height          =   315
         Left            =   1100
         TabIndex        =   3
         Top             =   1305
         Width           =   2760
      End
      Begin VB.ComboBox CboImpr 
         Height          =   315
         Left            =   1100
         TabIndex        =   0
         Top             =   225
         Width           =   2760
      End
      Begin VB.Image Image1 
         Height          =   690
         Left            =   4050
         Picture         =   "FrmImpresora.frx":030A
         Stretch         =   -1  'True
         Top             =   1125
         Width           =   780
      End
      Begin VB.Label LblTama 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "-1"
         Height          =   240
         Left            =   4050
         TabIndex        =   19
         Top             =   1665
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Label Label3 
         Caption         =   "Tamaño"
         Height          =   285
         Left            =   90
         TabIndex        =   18
         Top             =   1665
         Width           =   1050
      End
      Begin VB.Label LblImpr 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "-1"
         Height          =   240
         Left            =   4050
         TabIndex        =   17
         Top             =   1125
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Label LblOrie 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "-1"
         Height          =   240
         Left            =   4050
         TabIndex        =   16
         Top             =   1395
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Label Label6 
         Caption         =   "Orientación"
         Height          =   285
         Left            =   90
         TabIndex        =   15
         Top             =   1350
         Width           =   1050
      End
      Begin VB.Label LblPuer 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1095
         TabIndex        =   2
         Top             =   945
         Width           =   2760
      End
      Begin VB.Label Label4 
         Caption         =   "Puerto"
         Height          =   240
         Left            =   90
         TabIndex        =   14
         Top             =   990
         Width           =   960
      End
      Begin VB.Label LblCont 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1095
         TabIndex        =   1
         Top             =   585
         Width           =   2760
      End
      Begin VB.Label Label2 
         Caption         =   "Controlador"
         Height          =   240
         Left            =   90
         TabIndex        =   13
         Top             =   630
         Width           =   960
      End
      Begin VB.Label Label1 
         Caption         =   "Impresoras"
         Height          =   285
         Left            =   90
         TabIndex        =   8
         Top             =   270
         Width           =   1050
      End
   End
   Begin VB.Frame FDato 
      Height          =   150
      Left            =   315
      TabIndex        =   9
      Top             =   45
      Visible         =   0   'False
      Width           =   555
      Begin VB.ListBox LTama 
         Height          =   1620
         Left            =   90
         TabIndex        =   20
         Top             =   225
         Width           =   2265
      End
      Begin VB.ListBox LPuer 
         Height          =   1620
         Left            =   90
         TabIndex        =   12
         Top             =   225
         Width           =   2265
      End
      Begin VB.ListBox LCont 
         Height          =   1620
         Left            =   90
         TabIndex        =   11
         Top             =   225
         Width           =   2265
      End
      Begin VB.ListBox LImpr 
         Height          =   1620
         Left            =   90
         TabIndex        =   10
         Top             =   225
         Width           =   2265
      End
   End
End
Attribute VB_Name = "FrmImpresora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xImpr As Printer

Private Sub CboImpr_Change()
KeyAscii = 0
End Sub

Private Sub CboImpr_Click()
If CboImpr.ListIndex = -1 Then Exit Sub
If CboOrie.ListIndex = -1 Then CboOrie.ListIndex = 0
If CboTama.ListIndex = -1 Then CboTama.ListIndex = 0
If CboImpr.ListCount = 0 Then Exit Sub
LblCont.Caption = LCont.List(CboImpr.ListIndex)
LblPuer.Caption = LPuer.List(CboImpr.ListIndex)
LblImpr.Caption = Val(CboImpr.ListIndex)
End Sub

Private Sub CboImpr_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then CboOrie.SetFocus
KeyAscii = 0
End Sub

Private Sub CboOrie_Change()
KeyAscii = 0
End Sub

Private Sub CboOrie_Click()
If CboOrie.ListIndex = -1 Then Exit Sub
LblOrie.Caption = Val(CboOrie.ListIndex) + 1
End Sub

Private Sub CboOrie_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then CboTama.SetFocus
KeyAscii = 0
End Sub

Private Sub CboTama_Change()
KeyAscii = 0
End Sub

Private Sub CboTama_Click()
If CboTama.ListIndex = -1 Then Exit Sub
LblTama.Caption = LTama.List(CboTama.ListIndex)
End Sub

Private Sub CboTama_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then CmdAceptar.SetFocus
KeyAscii = 0
End Sub

Private Sub CmdAceptar_Click()
If Val(LblImpr.Caption) < 0 Then Exit Sub
xmImpresora = LImpr.List(LblImpr.Caption)
xmControlador = LCont.List(LblImpr.Caption)
xmPuerto = LPuer.List(LblImpr.Caption)
xmOrientacion = Val(LblOrie.Caption)
CmdCancelar_Click
End Sub

Private Sub CmdCancelar_Click()
Unload Me
ValoresIniciales
End Sub

Private Sub Form_Load()
ListarImpresora
End Sub

Public Sub ListarImpresora()
On Error Resume Next

'Lista de impresoras registradas
a = 0
For Each xImpr In Printers
    a = a + 1
    If a = 1 Then
    
    End If
    CboImpr.AddItem xImpr.DeviceName    'Nombre
    LImpr.AddItem xImpr.DeviceName      'Nombre
    LPuer.AddItem xImpr.Port            'Puerto
    LCont.AddItem xImpr.DriverName      'Controlador
Next
CboOrie.AddItem "Vertical"
CboOrie.AddItem "Horizontal"
CboTama.AddItem "Carta           (216 x 279)"
LTama.AddItem 1
CboTama.AddItem "Oficio           (216 x 279)"
LTama.AddItem 5
CboTama.AddItem "A3               (216 x 279)"
LTama.AddItem 8
CboTama.AddItem "A4               (216 x 279)"
LTama.AddItem 9
CboTama.AddItem "Contable      (310 x 280)"
LTama.AddItem 39
End Sub

Private Sub Form_Resize()
If Me.Height = 300 Or Me.Width = 2400 Then Exit Sub
Me.Height = 2535
Me.Width = 5220
End Sub

Public Sub ValoresIniciales()
On Error Resume Next

End Sub

