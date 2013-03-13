VERSION 5.00
Begin VB.Form frOpenGr 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Abrir archivo de imágenes"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5775
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   5775
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   3120
      TabIndex        =   4
      Top             =   4680
      Width           =   1575
   End
   Begin VB.CommandButton cmAbrir 
      Caption         =   "&Abrir"
      Height          =   375
      Left            =   1080
      TabIndex        =   3
      Top             =   4680
      Width           =   1575
   End
   Begin VB.FileListBox xFile 
      Height          =   4380
      Left            =   3360
      Pattern         =   "*.bmp;*.jpg;*.gif;*.dib;*.wmf;*.emf"
      TabIndex        =   2
      Top             =   120
      Width           =   2295
   End
   Begin VB.DirListBox xDir 
      Height          =   1665
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   3135
   End
   Begin VB.DriveListBox xDrive 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
   Begin VB.Image xFoto 
      Height          =   2295
      Left            =   120
      Top             =   2280
      Width           =   3135
   End
End
Attribute VB_Name = "frOpenGr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub CMABRIR_Click()
    If xFoto.Tag = "" Then
        MsgBox "Asignación de gráfico inválido"
        Exit Sub
    End If
    VGUTIL(0) = xFoto.Tag
    Unload Me
End Sub

Private Sub cmCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    VGUTIL(0) = ""
End Sub

Private Sub xDir_Change()
    xFile.PATH = xDir.PATH
End Sub

Private Sub xDrive_Change()
    On Error GoTo ErrDrive
    xDir.PATH = xDrive.Drive
    Exit Sub
ErrDrive:
    MsgBox "La unidad seleccionada no está lista", vbCritical
    Exit Sub
End Sub

Private Sub xFile_Click()
On Error Resume Next
    Dim RUTA As String
    If xFile.FileName <> "" Then
       If Right(xFile.PATH, 1) = "\" Then
          RUTA = xFile.PATH
         Else:
            RUTA = xFile.PATH & "\"
       End If
       xFoto.Picture = LoadPicture(RUTA & xFile.FileName)
       xFoto.Tag = RUTA & xFile.FileName
    End If
End Sub
