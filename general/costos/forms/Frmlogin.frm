VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmLogin 
   Caption         =   "Form1"
   ClientHeight    =   3180
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5415
   LinkTopic       =   "Form1"
   Picture         =   "Frmlogin.frx":0000
   ScaleHeight     =   3180
   ScaleWidth      =   5415
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Aceptar"
      Height          =   675
      Left            =   3900
      MaskColor       =   &H000000FF&
      Picture         =   "Frmlogin.frx":0342
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1035
      Width           =   775
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   675
      Left            =   3915
      Picture         =   "Frmlogin.frx":0784
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1755
      Width           =   775
   End
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   1980
      MaxLength       =   8
      TabIndex        =   2
      Top             =   1065
      Width           =   1725
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1980
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1665
      Width           =   1725
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      ItemData        =   "Frmlogin.frx":0BC6
      Left            =   1980
      List            =   "Frmlogin.frx":0BD0
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   615
      Visible         =   0   'False
      Width           =   2670
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1980
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   225
      Width           =   2655
   End
   Begin MSComCtl2.DTPicker DTPfecha 
      Height          =   375
      Left            =   1980
      TabIndex        =   4
      Top             =   2145
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   107937793
      CurrentDate     =   39104
   End
   Begin VB.Label lblLabels 
      Caption         =   "Có&digo de Usuario:"
      Height          =   270
      Index           =   0
      Left            =   420
      TabIndex        =   12
      Top             =   1065
      Width           =   1440
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Contraseña:"
      Height          =   270
      Index           =   1
      Left            =   405
      TabIndex        =   11
      Top             =   1665
      Width           =   1080
   End
   Begin VB.Label Label5 
      Caption         =   "Acceso      : "
      Height          =   300
      Left            =   420
      TabIndex        =   10
      Top             =   660
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Label Label4 
      Caption         =   "Empresa     :"
      Height          =   255
      Index           =   0
      Left            =   420
      TabIndex        =   9
      Top             =   225
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Version 2007.06.01"
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   405
      TabIndex        =   8
      Top             =   2730
      Width           =   3315
   End
   Begin VB.Label Label4 
      Caption         =   "Fecha de Ttrabajo   :"
      Height          =   255
      Index           =   1
      Left            =   420
      TabIndex        =   6
      Top             =   2265
      Width           =   1575
   End
End
Attribute VB_Name = "FrmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Reg1 As New ADODB.Recordset

Private Sub cmdCancel_Click()
If MDIPrincipal.Visible = True Then
    Unload Me
Else
    End
End If
End Sub

Private Sub cmdOK_Click()
Set VGvardllgen = New dllgeneral.dll_general
On Error GoTo Erla
 Call adicionarcampos
 Call adicionarcamposcostos
VGParamSistem.FechaTrabajo = DTPfecha.Value
VGParamSistem.Anoproceso = Format(Year(DTPfecha), "0")
VGParamSistem.Mesproceso = Format(Month(DTPfecha), "0")
 
If clave < 4 Then
    
    If Buscar_en_BD Then
        VGUsuario = txtUserName
        VGPass = txtPassword
        Dim Clsmenu As New ClassMenu
        Set Clsmenu.conexion = VGconfig
        VGtipo = Costos
        Clsmenu.TablaMenu = "si_menu"
        Clsmenu.CrearTablaMenu
        Clsmenu.TabaMenuDet = "si_menuusuarios"
        Clsmenu.TablaMenu = "si_menu"
        Call Clsmenu.HabilitarMenuNom(VGUsuario)
    Else
        MsgBox "La contraseña no es válida. Vuelva a intentarlo", , "Inicio de sesión"
        txtPassword.SetFocus
        SendKeys "{Home}+{End}"
        clave = clave + 1
        If clave < 3 Then
            If Combo3.ListIndex = 0 Then
                    MsgBox "Administrador No Autorizado", vbCritical, "Inicio de sesión"
            Else
                    MsgBox "Usuario No Autorizado", vbCritical, "Inicio de sesión"
            End If
        Else
            End
        End If
        Exit Sub
    End If
Else
    MsgBox "Usuario No Autorizado", vbCritical, "Inicio de sesión"
    
End If

MDIPrincipal.Show
    MDIPrincipal.StatusBar1.Panels(1).Text = "Mes Proceso : " & VGvardllgen.DesMes(Month(DTPfecha))
    MDIPrincipal.StatusBar1.Panels(2).Text = "Año Proceso : " & VGParamSistem.Anoproceso
    MDIPrincipal.StatusBar1.Panels(3).Text = "Fecha de Trabajo : " & VGParamSistem.FechaTrabajo & ""
    MDIPrincipal.StatusBar1.Panels(4).Text = "Tipo Cambio  : " & Format(tccambio, "#.000") & ""
    MDIPrincipal.StatusBar1.Panels(5).Text = "Servidor : " & VGParamSistem.Servidor
    MDIPrincipal.StatusBar1.Panels(4).Text = "Tipo Cambio  (" & Format(tccambio, "#.000") & ")"

Exit Sub
Erla:
Resume Next
        MsgBox Err.Description
End Sub

Public Function Buscar_en_BD() As Boolean
Set reg = New ADODB.Recordset

Combo3.ListIndex = 1
If Combo3.ListIndex = 0 Then 'Administrador
    
            reg.Open "Select * from Administrador", VGconfig, adOpenStatic
            
            If reg.RecordCount <> 0 Then
                reg.MoveFirst
                Do While Not reg.EOF
                    If UCase(Trim(txtUserName)) = UCase(reg.Fields("ADM_NOMBRE")) And UCase(RTrim$(txtPassword)) = DECODIFICA(reg.Fields("ADM_PASSWORD"), NUMMAGICO) Then
                            Buscar_en_BD = True
                            VGUsua = UCase(txtUserName)
                            VGPass = txtPassword
                            vGAdmLog = True
                            Exit Do
                    End If
                    reg.MoveNext
                    If reg.EOF Then Exit Do
                Loop
            End If
     'End If
ElseIf Combo3.ListIndex = 1 Then 'Usuario
    vGAdmLog = False
    reg.Open "Select * from si_USUARIO", VGconfig, adOpenStatic
     Buscar_en_BD = False
     If reg.RecordCount <> 0 Then
     reg.MoveFirst
     Do While Not reg.EOF
        If UCase(RTrim$(txtUserName)) = UCase(reg.Fields("USU_CODIGO")) And UCase(RTrim$(txtPassword)) = DECODIFICA(reg.Fields("USU_PASSWORD"), NUMMAGICO) Then
            VGUsua = UCase(txtUserName)
            VGPass = txtPassword
            Buscar_en_BD = True
            Exit Do
        End If
        reg.MoveNext
        If reg.EOF Then Exit Do
     Loop
    End If
End If
Set reg = Nothing
End Function

Private Sub Combo1_Click()
Dim sFileName As String
Dim sBD As String
Dim RSQL As String
Dim rs As ADODB.Recordset
Dim basect  As String
On Local Error GoTo ERRAR
basect = VGParamSistem.BDEmpresaCT
If Combo1.ListCount > 0 Then
    Set rs = New ADODB.Recordset
    rs.Open "Select * from EMPRESA WHERE  EMP_RAZON_NOMBRE = '" & Trim(Combo1.Text) & "' ", VGconfig, adOpenStatic
    If Not rs.EOF Then
      VGCODEMPRESA = rs("EMP_CODIGO")
      VGParametros.NomEmpresa = ESNULO(rs("EMP_RAZON_NOMBRE"), "")
      VGParametros.RucEmpresa = ESNULO(Reg1("EMP_RUC_DOCUMENTO"), "")
      If Trim(rs!empresabaseinventarios) <> "" Then
         VGParamSistem.BDEmpresa = rs!empresabaseinventarios
         Set VGCNx = New ADODB.Connection
         VGCNx.CursorLocation = adUseClient
         VGCNx.CommandTimeout = 0
         VGCNx.ConnectionTimeout = 0
         VGCNx.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & VGParamSistem.Usuario & ";Password=" & VGParamSistem.Pwd & ";Initial Catalog=" & VGParamSistem.BDEmpresa & ";Data Source=" & VGParamSistem.Servidor
         VGCNx.Open
         VGParamSistem.BDEmpresaCT = VGdllApi.LeerIni(App.Path & "\MARFICE.INI", "CONTABILIDAD", "BDDATOS", "?")
         If VGParamSistem.BDEmpresaCT = "" And basect <> VGParamSistem.BDEmpresa Then
            VGParamSistem.BDEmpresaCT = VGParamSistem.BDEmpresa
            Set VGcnxCT = New ADODB.Connection
            VGcnxCT.CursorLocation = adUseClient
            VGcnxCT.CommandTimeout = 0
            VGcnxCT.ConnectionTimeout = 0
            VGcnxCT.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & VGParamSistem.UsuarioCT & ";Password=" & VGParamSistem.PwdCT & ";Initial Catalog=" & VGParamSistem.BDEmpresaCT & ";Data Source=" & VGParamSistem.ServidorCT
            VGcnxCT.Open
         End If
      End If
    Else
      VGCODEMPRESA = Reg1("EMP_CODIGO")
      VGParametros.NomEmpresa = Reg1("EMP_RAZON_NOMBRE")
      'Fernando: 04/09/2001:
      VGParametros.RucEmpresa = Reg1("EMP_RUC_DOCUMENTO")
      '***
    End If
    rs.Close
    Call adicionarcampos
    RSQL = "Select  * From al_sistema"
    Set rs = New ADODB.Recordset
    Set rs = VGCNx.Execute(RSQL)
    If rs.RecordCount > 0 Then
    End If
    rs.Close
    
    MDIPrincipal.Caption = "Sistema de Costos Gerenciales " & "       " & VGParametros.RucEmpresa
    
End If
Exit Sub

ERRAR:
MsgBox "Ocurrio un Error," & error & " debe Actualizar su Base de Datos e Ingrese Nuevamente al Sistema", vbInformation, "Aviso.."
Resume Next
End Sub





Private Sub Form_Load()
 clave = 1
ADOCONECTAR
DTPfecha.Value = Date
Set Reg1 = New ADODB.Recordset
Reg1.Open "Select * from EMPRESA where empresaflaginventarios= 1 order by EMP_CODIGO ", VGconfig, adOpenStatic
LlenarListBox


If Combo1.ListCount > 0 Then Combo1.ListIndex = 0
If Combo3.ListCount > 0 Then Combo3.ListIndex = 0
End Sub


Private Sub txtPassword_GotFocus()
With txtPassword
 .SelStart = 0
 .SelLength = Len(.Text)
End With
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 SendKeys "{tab}"
 KeyAscii = 0
End If
End Sub

Private Sub txtUserName_GotFocus()
With txtUserName
 .SelStart = 0
 .SelLength = Len(.Text)
End With
End Sub

Private Sub txtUserName_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
    SendKeys "{tab}"
    KeyAscii = 0
End If
End Sub

Public Sub LlenarListBox()
 If Reg1.EOF Then Exit Sub
 If Reg1.BOF Then Exit Sub
 Do While Not Reg1.EOF
    If Not IsNull(Reg1.Fields("EMP_RAZON_NOMBRE")) Then
        Combo1.AddItem Reg1.Fields("EMP_RAZON_NOMBRE")
    End If
      Reg1.MoveNext
      If Reg1.EOF Then Exit Do
 Loop
 Reg1.MoveFirst
 Combo1.ListIndex = 0
End Sub


