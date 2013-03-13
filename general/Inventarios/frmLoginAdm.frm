VERSION 5.00
Begin VB.Form frmLoginAdm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Seguridad a Nivel de Administradores"
   ClientHeight    =   2850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6390
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   6390
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   240
      TabIndex        =   7
      Top             =   1800
      Width           =   5895
      Begin VB.CommandButton cmdContinuar 
         Caption         =   "C&ontinuar ..."
         Height          =   375
         Left            =   4320
         TabIndex        =   0
         Top             =   300
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Nota: Si no es Administrador pulse [Continuar]"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   4095
      End
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   2370
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1005
      Width           =   2325
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   675
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   960
      Width           =   775
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   675
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   775
   End
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   2370
      MaxLength       =   10
      TabIndex        =   1
      Top             =   375
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Contraseña:"
      Height          =   270
      Index           =   1
      Left            =   1185
      TabIndex        =   6
      Top             =   1020
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "Có&digo de Administrador:"
      Height          =   390
      Index           =   0
      Left            =   1185
      TabIndex        =   5
      Top             =   390
      Width           =   1080
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   240
      Stretch         =   -1  'True
      Top             =   480
      Width           =   735
   End
End
Attribute VB_Name = "frmLoginAdm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim cn As ADODB.Connection
Dim reg As ADODB.Recordset

Dim Clave As Integer
Private Sub cmdCancel_Click()
'establecer la variable global a false
'para indicar un inicio de sesión fallido
 VGAdmLogin = False
 End
End Sub

Private Sub cmdContinuar_Click()
 Unload Me
 If Not VGAdmLogin Then 'usuarios de empresa
  frmEmpresa.Show
 Else 'administrador
  frmPrincipal.StatusBar1.Panels.Item(4).text = "NINGUNA"
  frmPrincipal.StatusBar1.Panels.Item(6).text = VGFecTrb
  HabilitarMenu_Usuarios "A1"
 End If
End Sub

Private Sub cmdOK_Click()
 If Clave < 3 Then
    'comprobar si la contraseña es correcta
    If Buscar_en_BD() Then
        'colocar código aquí para pasar al sub
        'que llama si la contraseña es correcta
        VGAdmLogin = True
        VGADM_NOMBRE = Trim(txtUserName)
        VGADM_PASSWORD = Trim(txtPassword)
        cmdContinuar_Click
        Exit Sub
    Else
        MsgBox "La contraseña no es válida. Vuelva a intentarlo", , "Inicio de sesión"
        txtPassword.SetFocus
        SendKeys "{Home}+{End}"
        Clave = Clave + 1
        VGAdmLogin = False
    End If
 Else
    MsgBox "Admistrador no Registrado", vbCritical, "Seguridad a Nivel de Administradores"
    VGAdmLogin = False
    End
 End If
End Sub

Private Sub Form_Load()
 VGAdmLogin = False
 Set cn = New ADODB.Connection
 Set reg = New ADODB.Recordset
 reg.CursorType = adOpenDynamic
 With cn
  .Provider = "Microsoft.Jet.OLEDB.3.51"
  .ConnectionString = "Data Source=C:\WENCO\BDWENCO.MDB;Jet OLEDB:Database Password=segura;"
  .Open
 End With
 reg.ActiveConnection = cn
 reg.Open "select * from ADMINISTRADOR"
 Clave = 1
End Sub
Public Function Buscar_en_BD() As Boolean
 Buscar_en_BD = False
 reg.MoveFirst
 Do While Not reg.EOF
  If UCase(Trim(txtUserName)) = UCase(reg.Fields("ADM_NOMBRE")) And UCase(RTrim$(txtPassword)) = DECODIFICA(reg.Fields("ADM_PASSWORD"), NUMMAGICO) Then
   Buscar_en_BD = True
  End If
  reg.MoveNext
 Loop
End Function
