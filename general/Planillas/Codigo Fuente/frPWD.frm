VERSION 5.00
Begin VB.Form frPWD 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Contraseña de Ingreso"
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4635
   Icon            =   "frPWD.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   4635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   330
      Left            =   2460
      TabIndex        =   6
      Top             =   4620
      Width           =   1140
   End
   Begin VB.CommandButton cmAceptar 
      Caption         =   "&Aceptar"
      Height          =   330
      Left            =   1080
      TabIndex        =   5
      Top             =   4620
      Width           =   1140
   End
   Begin VB.Frame Frame2 
      Caption         =   "Identificación"
      Height          =   1890
      Left            =   135
      TabIndex        =   0
      Top             =   2625
      Width           =   4365
      Begin VB.ComboBox CmbAcc 
         Height          =   315
         ItemData        =   "frPWD.frx":0442
         Left            =   1995
         List            =   "frPWD.frx":044C
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   405
         Width           =   1920
      End
      Begin VB.TextBox xPass 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   2025
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   1410
         Width           =   1875
      End
      Begin VB.TextBox xUser 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   2025
         MaxLength       =   10
         TabIndex        =   2
         Top             =   1035
         Width           =   1875
      End
      Begin VB.Label Label3 
         Caption         =   "Acceso Como"
         Height          =   300
         Left            =   330
         TabIndex        =   7
         Top             =   435
         Width           =   1290
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Contraseña"
         Height          =   195
         Left            =   945
         TabIndex        =   3
         Top             =   1470
         Width           =   810
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   345
         Picture         =   "frPWD.frx":0468
         Top             =   1470
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Usuario:"
         Height          =   195
         Left            =   945
         TabIndex        =   1
         Top             =   1110
         Width           =   585
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   345
         Picture         =   "frPWD.frx":08AA
         Top             =   915
         Width           =   240
      End
   End
   Begin VB.Image Image3 
      BorderStyle     =   1  'Fixed Single
      Height          =   2490
      Left            =   135
      Top             =   75
      Width           =   4365
   End
End
Attribute VB_Name = "frPWD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim VRSALIR As Boolean

Private Sub CMACEPTAR_CLICK()
Dim CNW As New ADODB.Connection
Dim PWD As String
Dim NUSER As Integer
Dim X As Integer
    Call CLMENU.MenuTrue
    REGSISTEMA.USER = "INVITADO"
    If Len(Trim(Me.xPass.Text)) = 0 Then
        xPass.SetFocus
        Exit Sub
    End If
    If Len(Trim(Me.xUser.Text)) = 0 Then
        xUser.SetFocus
        Exit Sub
    End If
    With CNW
        .CursorLocation = adUseClient
        .CommandTimeout = 0
        .ConnectionTimeout = 0
        If VGL_INTEGRWNT Then
            .ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=" & VGL_BASE & ""
         Else
        '    .ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=" & VGL_BASE & ""
            .ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=SOPORTE;Password=SOPORTE;Initial Catalog=" & VGL_BASE & ";Data Source=" & VGL_SERVER
        End If
        .ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Password=;Initial Catalog=" & VGL_BASE & ";Data Source=" & VGL_SERVER
'       .ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;Password=sa;User ID=;Initial Catalog=" & VGL_BASE & ";Data Source=" & VGL_SERVER
        .Open
    End With
    
   
   'Verificar si existen Administradores
        CNW.Execute "UPDATE USUARIOS SET USUARIO=USUARIO", X
        If X = 0 Then
            PWD = PROCSIS.PrCodifica(Trim(xPass.Text), 5)
            CNW.Execute "INSERT INTO USUARIOS VALUES ('" & Me.xUser.Text & "','" & PWD & "','1111','A')"
        End If
    
    
    'EN CASO DE SER ADMINISTRADORES
    If CmbAcc.ListIndex = 0 Then
        'If Not (UCase(xUser.Text) = "STAR" And UCase(xPass.Text) = "CONSUL") Then
           If Not VERIFICAADMI(NUSER) Then Exit Sub
          Else: REGSISTEMA.USER = xUser.Text
        'End If
        If NUSER > 0 Then REGSISTEMA.USER = xUser.Text
        REGSISTEMA.PASSWORD = xPass.Text
        REGSISTEMA.ESADMINISTRADOR = True
    End If
    'EN CASO QUE SEA USUARIO
    If CmbAcc.ListIndex = 1 Then
        If Not VERIFICAUSUARIO(NUSER) Then Exit Sub
            If NUSER > 0 Then
                REGSISTEMA.USER = xUser.Text
                REGSISTEMA.PASSWORD = xPass.Text
                REGSISTEMA.ESADMINISTRADOR = False
                Call CLMENU.HabilitarMenuNom(Trim(xUser.Text), REGSISTEMA.RUC)
            End If
    End If
    MDIPrincipal.BarraEstado.Panels(1).Text = REGSISTEMA.USER & " "
    Unload Me
    Unload frPanEmp
End Sub
Private Function VERIFICAADMI(Optional ByRef NUMAD As Integer) As Boolean
    Dim RSPASS As New ADODB.Recordset
    Dim PWD As String
    VERIFICAADMI = False
    RSPASS.Open "USUARIOS", DBSTARPLAN, adOpenKeyset
    NUMAD = RSPASS.RecordCount
    If RSPASS.RecordCount = 0 Then
        VERIFICAADMI = True
        Exit Function
    End If
    'VALIDANDO SI EXISTE EL USUARIO
    Set RSPASS = Nothing
    RSPASS.Open "SELECT * FROM USUARIOS  WHERE USUARIO='" & xUser.Text & "'", DBSTARPLAN, adOpenKeyset
    If RSPASS.RecordCount = 0 Then
        MsgBox "NO SE ENCUENTRA EL ADMINISTRADOR", vbExclamation
        xUser.SetFocus
        Exit Function
    End If
    'VALIDANDO SI EXISTE EL PWD
    PWD = PROCSIS.PrCodifica(xPass.Text, 5)
    Set RSPASS = Nothing
    RSPASS.Open "SELECT * FROM USUARIOS WHERE USUARIO='" & xUser.Text & "' AND CLAVE='" & PWD & "'", DBSTARPLAN, adOpenKeyset
    If RSPASS.RecordCount = 0 Then
        MsgBox "LA CONTRASEÑA ES INCORRECTA", vbExclamation
        xPass.SetFocus
        Exit Function
    End If
    VERIFICAADMI = True
End Function
Private Function VERIFICAUSUARIO(Optional ByRef CANTUSERXEMP As Integer) As Boolean
    Dim RSPASS As New ADODB.Recordset
    Dim PWD As String
    CLMENU.TablaUsu = "USUARIO"
    VERIFICAUSUARIO = False
    RSPASS.Open "SELECT * FROM " & UCase(CLMENU.TablaUsu) & " WHERE EMP_CODIGO='" & Trim(REGSISTEMA.RUC) & "'", DBSTARPLAN, adOpenKeyset
    CANTUSERXEMP = RSPASS.RecordCount
    If RSPASS.RecordCount = 0 Then
        VERIFICAUSUARIO = True
        Exit Function
    End If
    'VALIDANDO SI EXISTE EL USUARIO
    Set RSPASS = Nothing
    RSPASS.Open "SELECT * FROM " & UCase(CLMENU.TablaUsu) & " WHERE USU_CODIGO='" & xUser.Text & _
    "' AND  LTRIM(EMP_CODIGO)='" & Trim(REGSISTEMA.RUC) & "'", DBSTARPLAN, adOpenKeyset
    If RSPASS.RecordCount = 0 Then
        MsgBox "NO SE ENCUENTRA EL USUARIO EN ESTA EMPRESA", vbExclamation
        xUser.SetFocus
        Exit Function
    End If
    
    'VALIDANDO SI EXISTE EL PWD
    PWD = PROCSIS.PrCodifica(xPass.Text, 5)
    Set RSPASS = Nothing
    RSPASS.Open "SELECT * FROM " & UCase(CLMENU.TablaUsu) & " WHERE USU_CODIGO='" & xUser.Text & _
    "' AND USU_PASSWORD='" & PWD & "' AND LTRIM(EMP_CODIGO)='" & Trim(REGSISTEMA.RUC) & "'", DBSTARPLAN, adOpenKeyset
    If RSPASS.RecordCount = 0 Then
        MsgBox "LA CONTRASEÑA ES INCORRECTA", vbExclamation
        xPass.SetFocus
        Exit Function
    End If
    VERIFICAUSUARIO = True
End Function
Private Sub CMCANCELAR_CLICK()
    Unload Me
End Sub

Private Sub Form_Load()
    CmbAcc.ListIndex = 0
    If UCase(Dir$(REGSISTEMA.PATH & "\LOGINMGN.JPG")) = "LOGINMGN.JPG" Then Image3.Picture = LoadPicture(REGSISTEMA.PATH & "\LOGINMGN.JPG")
End Sub

Private Sub xPass_GotFocus()
xPass.SelStart = 0
xPass.SelLength = Len(xPass.Text)
End Sub

Private Sub XPASS_KeyDown(KEYCODE As Integer, Shift As Integer)
If KEYCODE = 13 Then cmAceptar.SetFocus
End Sub

Private Sub xUser_GotFocus()
xUser.SelStart = 0
xUser.SelLength = Len(xUser.Text)
End Sub

Private Sub XUSER_KeyDown(KEYCODE As Integer, Shift As Integer)
    If KEYCODE = 13 Then xPass.SetFocus
End Sub


