VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inicio de Sesión"
   ClientHeight    =   2445
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   6465
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1444.587
   ScaleMode       =   0  'User
   ScaleWidth      =   6070.286
   StartUpPosition =   2  'CenterScreen
   Begin MSMask.MaskEdBox FechaBox 
      Height          =   315
      Left            =   2640
      TabIndex        =   4
      Top             =   1560
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   675
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   360
      Width           =   775
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   675
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1200
      Width           =   775
   End
   Begin VB.CommandButton cmdAntes 
      Caption         =   "<<Anterior"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1920
      Width           =   975
   End
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   2640
      MaxLength       =   5
      TabIndex        =   1
      Top             =   345
      Width           =   2325
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   2610
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   975
      Width           =   2325
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha de Trabajo :"
      Height          =   255
      Left            =   1200
      TabIndex        =   9
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "dd/mm/yyyy"
      Height          =   255
      Left            =   3840
      TabIndex        =   8
      Top             =   1560
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   360
      Stretch         =   -1  'True
      Top             =   480
      Width           =   855
   End
   Begin VB.Label lblLabels 
      Caption         =   "Có&digo de Usuario:"
      Height          =   390
      Index           =   0
      Left            =   1545
      TabIndex        =   0
      Top             =   360
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Contraseña:"
      Height          =   270
      Index           =   1
      Left            =   1545
      TabIndex        =   2
      Top             =   990
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim cn As ADODB.Connection
Dim reg As ADODB.Recordset

Dim DB As ADODB.Connection
Dim AdoReg1 As ADODB.Recordset


Dim Clave As Integer

Private Sub cmdAntes_Click()
 frmEmpresa.Show
 Unload Me
End Sub

Private Sub cmdCancel_Click()
    'establecer la variable global a false
    'para indicar un inicio de sesión fallido
    MsgBox "Ud. Saldrá del Sistema de Contabilidad", vbInformation, "Sistema de Contabilidad"
    Unload Me
    End
End Sub

Private Sub cmdOK_Click()
 If Clave < 3 Then
    'comprobar si la contraseña es correcta
    If Buscar_en_BD() Then
        If Busca_Fecha() Then
         HabilitarMenu_Usuarios VGUSU_NIVEL
         frmPrincipal.StatusBar1.Panels.Item(6).text = VGFecTrb
         Unload Me
        End If
    Else
        MsgBox "La contraseña no es válida. Vuelva a intentarlo", , "Inicio de sesión"
        txtPassword.SetFocus
        SendKeys "{Home}+{End}"
        Clave = Clave + 1
    End If
 Else
    MsgBox "Usuario No Autorizado", vbCritical, "Inicio de sesión"
    End
 End If
End Sub

Private Sub FechaBox_GotFocus()
With FechaBox
 .SelStart = 0
 .SelLength = .MaxLength
End With
End Sub

Private Sub Form_Load()
 frmLogin.Caption = frmLogin.Caption & " - " & VGEMP_RAZON
 Clave = 1
 Set cn = New ADODB.Connection
 Set reg = New ADODB.Recordset
 reg.CursorType = adOpenDynamic
 With cn
  .Provider = "Microsoft.Jet.OLEDB.3.51"
  .ConnectionString = "Data Source=C:\WENCO\BDWENCO.MDB;Jet OLEDB:Database Password=segura;"
  .Open
 End With
 reg.ActiveConnection = cn
 reg.Open "select * from USUARIO"
 ADOConectar
 FechaBox.text = VGFecTrb
End Sub

Public Function Buscar_en_BD() As Boolean
 Buscar_en_BD = False
 reg.MoveFirst
 Do While Not reg.EOF
  If UCase(RTrim$(txtUserName)) = UCase(reg.Fields("USU_CODIGO")) And UCase(RTrim$(txtPassword)) = DECODIFICA(reg.Fields("USU_PASSWORD"), NUMMAGICO) Then
   VGUSU_NIVEL = reg.Fields("USU_NIVEL")
   Buscar_en_BD = True
  End If
  reg.MoveNext
 Loop
End Function

Private Sub txtPassword_GotFocus()
With txtPassword
 .SelStart = 0
 .SelLength = Len(.text)
End With
End Sub

Private Sub txtUserName_GotFocus()
With txtUserName
 .SelStart = 0
 .SelLength = Len(.text)
End With
End Sub
Private Function Busca_Fecha() As Boolean
Dim Flag As Boolean
Flag = False
If ValidFecha(FechaBox) Then
 With AdoReg1
 If .RecordCount <> 0 Then
  .MoveFirst
  Do While Not .EOF
   If Right(FechaBox.text, 4) = .Fields(0) Then
    Flag = True
    Exit Do
   End If
   .MoveNext
  Loop
 End If
 End With
 If Not Flag Then
  MsgBox "No existe ejercicio contable para esta empresa", vbInformation, "Ingreso de Datos"
  FechaBox.SetFocus
 Else
 FechaBox.PromptInclude = True
 VGFecTrb = FechaBox.text
 FechaBox.PromptInclude = False
 End If
Else
 MsgBox "Verifique la Fecha que ingreso. No es validad", vbInformation, "Ingreso de Fecha"
 FechaBox.SetFocus
End If
Busca_Fecha = Flag
End Function

Public Sub ADOConectar()
 Set DB = New ADODB.Connection
 Set AdoReg1 = New ADODB.Recordset
 
 AdoReg1.CursorType = adOpenDynamic
 With DB
  .CursorLocation = adUseClient
  .Provider = "Microsoft.Jet.OLEDB.3.51"
  .ConnectionString = "Data Source=" & WENCOPATH & VGEMP_CODIGO & "\" & NAMEBD
  .Open
 End With
 AdoReg1.Open "Select * from EJERCICIO_CONTABLE", DB, adOpenStatic, adLockOptimistic
End Sub

