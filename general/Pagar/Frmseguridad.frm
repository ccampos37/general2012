VERSION 5.00
Begin VB.Form Frmseguridad 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1410
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   3705
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1410
   ScaleWidth      =   3705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame7 
      Height          =   1245
      Left            =   120
      TabIndex        =   3
      Top             =   60
      Width           =   3345
      Begin VB.TextBox Text2 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1860
         MaxLength       =   4
         PasswordChar    =   "*"
         TabIndex        =   0
         Top             =   270
         Width           =   945
      End
      Begin VB.CommandButton cAcepta 
         Caption         =   "Acepta"
         Height          =   375
         Left            =   570
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   780
         Width           =   1065
      End
      Begin VB.CommandButton cCancela 
         Caption         =   "Cancela"
         Height          =   375
         Left            =   1830
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   780
         Width           =   1065
      End
      Begin VB.Label Label4 
         Caption         =   "Seguridad"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   840
         TabIndex        =   4
         Top             =   330
         Width           =   1215
      End
   End
End
Attribute VB_Name = "Frmseguridad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cAcepta_Click()
  Dim nnume As Double
  Dim opera As Double
  Dim J As Integer
  Dim resi As Integer
  
  If Len(Trim$(Text2)) = 0 Then
     Exit Sub
  End If
  
  opera = 0
  For J = 1 To Len(Trim$(Text2))
     nnume = Mid$(Text2, J, 1)
     If J = 2 Then
       opera = opera * nnume
     Else
      opera = opera + nnume
     End If
  Next J
  resi = CDbl(Text2) Mod 2
  If resi <> 0 Then
     opera = opera + 1
  End If
  If opera = Day(FrmPlanillaCobranzaModi.MBox1) Then
     nAyuda = "1"
     Unload Me
  Else
     
     MsgBox "La contraseña no es valida...!!!", vbInformation, MsgTitle
     VGCNx.Execute "insert into sysseguridad values ('" & Date & "','" & Time & "','" & VGusuario & "'," & _
                "'" & "Contraseña no valida : " & Text2 & " ==>> cuando intento ingresar a eliminar documento de planilla de cobranza' )"
     
  End If
End Sub

Private Sub cCancela_Click()
  nAyuda = "0"
  Unload Me
End Sub

Private Sub Form_Load()
  MostrarForm Me, "C2"
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
  Call Seguir(Text2, KeyAscii)
End Sub
