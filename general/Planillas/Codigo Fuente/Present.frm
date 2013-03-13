VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Present 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4125
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6810
   Icon            =   "Present.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4125
   ScaleWidth      =   6810
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   5385
      Top             =   645
   End
   Begin MSComctlLib.ProgressBar Prog 
      Height          =   105
      Left            =   60
      TabIndex        =   1
      Top             =   45
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   185
      _Version        =   393216
      Appearance      =   0
      Max             =   50
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   3000
      X2              =   4260
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Versión SQL 200"
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
      Height          =   195
      Index           =   0
      Left            =   2400
      TabIndex        =   0
      Top             =   3120
      Width           =   1440
   End
End
Attribute VB_Name = "Present"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
    Dim N As Integer
    A = "" & Date
    If Not IsNumeric(Left(Right(A, 3), 1)) Then
        Timer1.Enabled = False
        MsgBox "El formato de fecha de su computadora no es válido.Use el formato: dd/mm/aaaa" & Chr(13) & Chr(10) & "Cambie el formato desde el Panel de Control, en la opción Configuración Regional", vbInformation
        Print Shell("CONTROL", vbNormalFocus)
        End
    End If
    A = Format("10.02", "0.00")
    If InStr(A, ",") > 0 Then
        Timer1.Enabled = False
        MsgBox "El formato de número de su computadora no es válido. Use como Símbolo decimal el punto y separador de miles a la coma. Cambie el formato desde el Panel de Control, en la opción Configuración Regional", vbInformation
        Print Shell("CONTROL", vbNormalFocus)
        End
    End If
End Sub

Private Sub Form_Load()
'    Me.Picture = LoadResPicture(101, 0)
    Timer1.Enabled = True
    Prog.Value = 0
    Prog.Max = 1000
End Sub

Private Sub Timer1_Timer()
    If Prog.Value >= Prog.Max - 10 Then Unload Me
    Prog.Value = Prog.Value + 30
End Sub

