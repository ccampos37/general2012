VERSION 5.00
Begin VB.Form frmlogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inicio de Sesion"
   ClientHeight    =   3270
   ClientLeft      =   2580
   ClientTop       =   2655
   ClientWidth     =   6495
   ControlBox      =   0   'False
   Icon            =   "frmloginIni.frx":0000
   LinkTopic       =   "Form11"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   6495
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2640
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   600
      Width           =   2655
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      ItemData        =   "frmloginIni.frx":08CA
      Left            =   2640
      List            =   "frmloginIni.frx":08D4
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   990
      Visible         =   0   'False
      Width           =   2670
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   2640
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2400
      Width           =   2325
   End
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   2640
      MaxLength       =   8
      TabIndex        =   0
      Top             =   1800
      Width           =   2325
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   675
      Left            =   5292
      Picture         =   "frmloginIni.frx":08F0
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2496
      Width           =   775
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Aceptar"
      Height          =   675
      Left            =   5280
      Picture         =   "frmloginIni.frx":0D32
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1776
      Width           =   775
   End
   Begin VB.Label Label1 
      Caption         =   "Version 2003.09.02"
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   105
      TabIndex        =   11
      Top             =   2985
      Width           =   3315
   End
   Begin VB.Label Label4 
      Caption         =   "Empresa     :"
      Height          =   255
      Left            =   1440
      TabIndex        =   10
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "Acceso      : "
      Height          =   300
      Left            =   1440
      TabIndex        =   9
      Top             =   1035
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Contraseña:"
      Height          =   270
      Index           =   1
      Left            =   1425
      TabIndex        =   6
      Top             =   2400
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "Có&digo de Usuario:"
      Height          =   390
      Index           =   0
      Left            =   1440
      TabIndex        =   5
      Top             =   1800
      Width           =   1080
   End
   Begin VB.Label Label3 
      Caption         =   "Control de  Inventarios"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   1440
      TabIndex        =   4
      Top             =   120
      Width           =   3975
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   6360
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      X1              =   120
      X2              =   6360
      Y1              =   1455
      Y2              =   1455
   End
End
Attribute VB_Name = "frmlogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim reg As ADODB.Recordset
Dim Reg1 As ADODB.Recordset
Dim clave As Integer

Public Sub ADOConectar()

cn.CursorLocation = adUseClient
cn.CommandTimeout = 0
cn.ConnectionTimeout = 200
cn.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID='" & Trim(VGBUsuario) & "';password='" & Trim(VGPassw) & "';Initial Catalog='" & Trim(VGBase2) & "';Data Source='" & Trim(VGServer2) & "'"
cn.Open

cRuta6 = sName & "" & cNomBd6          'BD. ConfigFac
Set cConexConf = New ADODB.Connection  'BD. ConfigFac
cConexConf.CursorLocation = adUseClient
'cConexConf.ConnectionString = "PROVIDER=Microsoft.Jet.OLEDB.3.51;Data Source=" & cRuta6 & ";Jet OLEDB:Database Password=segura;"
'cConexConf.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=bdwenco;Data Source=desarrollo2"
cConexConf.CommandTimeout = 0
cConexConf.ConnectionTimeout = 200
cConexConf.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & VGBUsuario & ";password='" & Trim(VGPassw) & "';Initial Catalog=bdwenco;Data Source=" & VGServer
cConexConf.Open

Set Reg1 = New ADODB.Recordset
'Reg1.Open "Select * from EMPRESA order by EMP_CODIGO ", cConexConf, adOpenStatic
Reg1.Open "Select * from EMPRESA order by EMP_CODIGO ", cn, adOpenStatic
LlenarListBox
End Sub

Private Sub cmdCancel_Click()
Dim op
    op = MsgBox("Ud. Saldrá del Sistema de Control de Inventario", vbInformation + vbYesNo, mensaje1)
    If op = vbYes Then
       cConexConf.Close
       'cn.Close
       Unload Me
       VGSALIR = True
     End If
End Sub

Private Sub cmdOK_Click()
On Error GoTo Erla
If clave < 4 Then
    
    If Buscar_en_BD Then
        If vGAdmLog Then
            'MDIMenu.StatusBar1.Panels(2).text = "Administrador: " & vGUsuario
            If Combo1.ListCount = 0 Then VGCOMP = ""
            'HabilitarMenu_Usuarios txtUserName, VGCOMP, "A1"
        Else
           FrmPrincipal.Men_SisPar.Enabled = False
           FrmPrincipal.Men_SisPar.Visible = False
        End If
        VGUsua = txtUserName
        VGPass = txtPassword
        'Ingresa Datos a la Tabla Menu
        Dim Clsmenu As New ClassMenu
        Set Clsmenu.Conexion = cConexConf
        Set Clsmenu.MDIMenu = FrmPrincipal
        Clsmenu.TablaMenu = "MENU_INV"
        Clsmenu.CrearTablaMenu
        
        Clsmenu.TabaMenuDet = "MEN_USU_INV"
        Clsmenu.TablaMenu = "MENU_INV"
        Call Clsmenu.HabilitarMenuNom(VGUsua, VGCOMP)
        
        
        Unload Me
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
    End
End If
Unload Me
Exit Sub
Erla:
Resume Next
        MsgBox Err.Description
End Sub

Public Function Buscar_en_BD() As Boolean
Set reg = New ADODB.Recordset
Combo3.ListIndex = 1
If Combo3.ListIndex = 0 Then 'Administrador
    
            reg.Open "Select * from Administrador", cConexConf, adOpenStatic
            
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
    reg.Open "Select * from USUARIO_INV Where EMP_CODIGO='" & VGCOMP & "'", cConexConf, adOpenStatic
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
Dim rsql As String
Dim rs As ADODB.Recordset
Dim Alma_defe As String
On Local Error GoTo ERRAR
If Combo1.ListCount > 0 Then
    Set rs = New ADODB.Recordset
    rs.Open "Select * from EMPRESA WHERE  EMP_RAZON_NOMBRE = '" & Trim(Combo1.text) & "' ", cConexConf, adOpenStatic
    If Not rs.EOF Then
      VGCOMP = rs("EMP_CODIGO")
      VGNemp = rs("EMP_RAZON_NOMBRE")
      'Fernando: 04/09/2001:
      VGRUCEMP = Reg1("EMP_RUC_DOCUMENTO")
      '***
    Else
      VGCOMP = Reg1("EMP_CODIGO")
      VGNemp = Reg1("EMP_RAZON_NOMBRE")
      'Fernando: 04/09/2001:
      VGRUCEMP = Reg1("EMP_RUC_DOCUMENTO")
      '***
    End If
    rs.Close
    cRutPath = sName & "DATA\" & VGCOMP & "\"
    cNomBd2 = "BDComun.mdb"
    cNomBd4 = VGNameCont & ".MDB"                                                '"BdContabilidad.mdb"
    cRuta5 = cRutPath & cNomBd2      'BD. Común
    cRuta2 = cRutPath & cNomBd2      'BD. Común
    cRuta3 = sName & cNomBd2          'BD. Común
    cRuta4 = cRutPath & cNomBd4      'BD. Contabilidad
 
    
    sFileName = App.Path & "\Inventario.ini"
    sBD = sGetIni(sFileName, "CONFIG", "BD", "?")
    If sBD <> "?" Then VGNameCont = sBD
    Set cConexCom = New ADODB.Connection  'BD. Común
    cConexCom.CursorLocation = adUseClient
    cConexCom.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & VGBUsuario & ";password='" & Trim(VGPassw) & "';Initial Catalog=" & VGBase & ";Data Source=" & VGServer
    cConexCom.Open
    
    If UCase(Dir$(cRuta4)) = UCase(cNomBd4) Then
        Set cConexCont = New ADODB.Connection  'BD. Contabilidad
        cConexCont.CursorLocation = adUseClient
        cConexCont.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & VGBUsuario & ";password='" & Trim(VGPassw) & "';Initial Catalog=" & cNomBd4 & ";Data Source=" & VGServer
        cConexCont.Open
    End If
    
    If Not ExisteElem(1, cConexCom, "CONFIGURACION", "Alma_defa") Then
       cConexCom.Execute "ALTER TABLE CONFIGURACION  ADD COLUMN  Alma_defa Text(10) "
    End If
    
    If Not ExisteElem(1, cConexCom, "CONFIGURACION", "TIPO_ALMA") Then ' V=ALMACEN VENTAS O S=ALMACEN SUMINISTROS
       cConexCom.Execute "ALTER TABLE CONFIGURACION  ADD COLUMN  TIPO_ALMA Text(1) "
       cConexCom.Execute "UPDATE CONFIGURACION SET TIPO_ALMA='V'"
    End If
    
    
    rsql = "select conf_codigo,alma_defa,tipo_alma,Ladrillera from configuracion"
    Set rs = New ADODB.Recordset
    rs.Open rsql, cConexCom, adOpenStatic
    If rs.RecordCount = 0 Then
        VGIASA = " "
    Else
        VGIASA = IIf(IsNull(rs("conf_codigo")), " ", rs("conf_codigo"))
        Alma_defe = cNull(rs!Alma_defa)
        VGTip_Alma = rs!tipo_Alma
        VGLadrillera = IIf(cNull(rs!Ladrillera) = "S", True, False)
    End If
    rs.Close
    
    If Alma_defe <> "" Then
       rsql = "Select  * From  TabAlm where taalma='" & Alma_defe & "'"
    Else
       rsql = "Select  * From  TabAlm"
    End If
    
    Set rs = New ADODB.Recordset
    rs.Open rsql, cConexCom, adOpenStatic
    If rs.RecordCount = 0 Then
        VGNomAlm = " "
    Else
        VGNomAlm = IIf(IsNull(rs("tadescri")), " ", rs("tadescri"))
        VGAlma = rs("taalma")
    End If
    rs.Close
    Set rs = cConexCom.Execute("select stockcomp from vt_parametroventa")
    stockcomp = rs.Fields(0)
    rs.Close
    FrmPrincipal.Caption = "Sistema de Inventario" & "     " & VGNomAlm & "    " & VGNemp
    
End If
Exit Sub

ERRAR:
MsgBox "Ocurrio un Error, debe Actualizar su Base de Datos e Ingrese Nuevamente al Sistema", vbInformation, "Aviso.."
Resume Next
End Sub

Private Sub Form_Load()
 Me.Cls
 clave = 1
central Me
ADOConectar
If Combo1.ListCount > 0 Then Combo1.ListIndex = 0
If Combo3.ListCount > 0 Then Combo3.ListIndex = 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
        VG_FecTrab = Format(Now, "dd/mm/yyyy")
End Sub

Private Sub txtPassword_GotFocus()
With txtPassword
 .SelStart = 0
 .SelLength = Len(.text)
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
 .SelLength = Len(.text)
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

