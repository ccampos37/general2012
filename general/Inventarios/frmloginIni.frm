VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmlogin 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   Caption         =   "Control de Almacenes"
   ClientHeight    =   5895
   ClientLeft      =   2535
   ClientTop       =   2250
   ClientWidth     =   10815
   FillColor       =   &H00E0E0E0&
   ForeColor       =   &H00E0E0E0&
   Icon            =   "frmloginIni.frx":0000
   LinkTopic       =   "Form11"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   10815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture 
      Height          =   5535
      Left            =   120
      Picture         =   "frmloginIni.frx":15DBA
      ScaleHeight     =   5475
      ScaleWidth      =   6075
      TabIndex        =   20
      Top             =   120
      Width           =   6135
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Height          =   5535
      Left            =   6360
      TabIndex        =   0
      Top             =   120
      Width           =   4215
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   0
         TabIndex        =   17
         Top             =   120
         Width           =   4215
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Control de Almacenes"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   600
            TabIndex        =   18
            Top             =   120
            Width           =   3060
         End
      End
      Begin VB.CommandButton cmdOK 
         BackColor       =   &H8000000E&
         Caption         =   "&Aceptar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1035
         Left            =   1110
         MaskColor       =   &H000000FF&
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   4050
         Width           =   1110
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H8000000E&
         Cancel          =   -1  'True
         Caption         =   "&Cancelar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1035
         Left            =   2340
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   4050
         Width           =   1110
      End
      Begin VB.TextBox txtUserName 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1410
         MaxLength       =   8
         TabIndex        =   4
         Top             =   2580
         Width           =   1725
      End
      Begin VB.TextBox txtPassword 
         Appearance      =   0  'Flat
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1410
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   3090
         Width           =   1725
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         ItemData        =   "frmloginIni.frx":3BE51
         Left            =   1380
         List            =   "frmloginIni.frx":3BE5B
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   5580
         Visible         =   0   'False
         Width           =   2670
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1410
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1020
         Width           =   2685
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1410
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1530
         Visible         =   0   'False
         Width           =   2700
      End
      Begin VB.ComboBox Combo4 
         ForeColor       =   &H00800000&
         Height          =   315
         ItemData        =   "frmloginIni.frx":3BE77
         Left            =   1410
         List            =   "frmloginIni.frx":3BE79
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   2040
         Width           =   2700
      End
      Begin MSComCtl2.DTPicker DTPfecha 
         Height          =   285
         Left            =   1410
         TabIndex        =   6
         Top             =   3600
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         Format          =   97386497
         CurrentDate     =   39104
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   19
         Top             =   5280
         Width           =   1605
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario :"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   300
         TabIndex        =   16
         Top             =   2640
         Width           =   630
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Contraseña:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   300
         TabIndex        =   15
         Top             =   3150
         Width           =   855
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Acceso      : "
         Height          =   300
         Left            =   540
         TabIndex        =   14
         Top             =   5655
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Grupo :"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   300
         TabIndex        =   13
         Top             =   1050
         Width           =   525
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha :"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   300
         TabIndex        =   12
         Top             =   3660
         Width           =   540
      End
      Begin VB.Label Lbempresa 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Empresa :"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   300
         TabIndex        =   11
         Top             =   1575
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pto.  Venta :"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   3
         Left            =   300
         TabIndex        =   10
         Top             =   2085
         Width           =   885
      End
   End
End
Attribute VB_Name = "frmlogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim reg As ADODB.Recordset
Dim REG1 As ADODB.Recordset
Dim clave As Integer



Sub LlenarListaEmpresas()
Dim REG1 As ADODB.Recordset

Set REG1 = VGCNx.Execute("Select * from co_multiempresas where empresacodigo<>'00'")
Combo2.Clear
Do While Not REG1.EOF
   Combo2.AddItem REG1.Fields("empresacodigo") + " " + REG1.Fields("empresadescripcion")
   REG1.MoveNext
Loop

Combo2.ListIndex = 0

End Sub

Private Sub cmdCancel_Click()
Dim op
    op = MsgBox("Esta seguro de salir del Sistema ?", vbQuestion + vbYesNo, "Inventarios")
    If op = vbYes Then
       VGConfig.Close
       Unload Me
       VGSALIR = True
     End If
End Sub

Private Sub cmdOK_Click()
Dim RSAUX  As New ADODB.Recordset
Dim RsPvta As New ADODB.Recordset, RsNemp As New ADODB.Recordset, RsUsu As New ADODB.Recordset
On Error GoTo Erla

If Len(Trim(txtUserName.text)) = 0 Then
    MsgBox "Usuario no valido. Vuelva a intentarlo", vbInformation, "Inicio de sesión"
    txtUserName.SetFocus
    Exit Sub
End If

If Len(Trim(txtPassword.text)) = 0 Then
    MsgBox "Contraseña no valida. Vuelva a intentarlo", vbInformation, "Inicio de sesión"
    txtPassword.SetFocus
    Exit Sub
End If
If DTPfecha > Date Then
    MsgBox "Fecha mayor al del sistema ", vbInformation, "Inicio de sesión"
    DTPfecha.SetFocus
    Exit Sub
End If


Call CargarParametrosCompras

If VGParametros.sistemamultiempresas Then
 VGParametros.empresacodigo = Left(Combo2.text, 2)
 VGParametros.NomEmpresa = Right(Combo2.text, Len(Combo2.text) - 2)
Else
 VGParametros.empresacodigo = "01"
End If
VGParametros.mesproceso = Format(DTPfecha, "yyyy") + Format(DTPfecha, "mm")

Set RsPvta = VGCNx.Execute("select puntovtacodigo,puntovtadescripcion from vt_puntoventa where puntovtacodigo='" & Left(Combo4.text, 2) & "'")
If RsPvta.RecordCount > 0 Then VGParametros.puntovta = RsPvta!puntovtacodigo

Set RsNemp = VGCNx.Execute("select empresadescripcion,empresaruc from co_multiempresas where empresacodigo='" & VGParametros.empresacodigo & "'")
If RsNemp.RecordCount > 0 Then VGParametros.NomEmpresa = RsNemp!empresadescripcion
If RsNemp.RecordCount > 0 Then VGParametros.RucEmpresa = RsNemp!empresaruc
   
VGParamSistem.fechatrabajo = DTPfecha.Value
VGParamSistem.AnoProceso = Format(Year(DTPfecha), "0")
VGParamSistem.mesproceso = Format(Month(DTPfecha), "0")

VGParamSistem.TablaCabcomprob = "co_cabeceraprovisiones"
VGParamSistem.tabladetcomprob = "co_detalleprovisiones"

Set RSAUX = New ADODB.Recordset
Set RSAUX = VGCNx.Execute(" select * from al_sistema where empresacodigo='" & Left(Combo2.text, 2) & "'")
If RSAUX.RecordCount > 0 Then
   VGflagconversioncodigo = ESNULO(RSAUX!flagconversioncodigo, 0)
   VGParametros.tipodevalorizacion = RSAUX!tipodevalorizacion
   VGParametros.SaldosvalorxAlmacen = RSAUX!SaldosvalorxAlmacen
   VGParametros.Valorestadooccodigo = RSAUX!Valorestadooccodigo
   VGParametros.SaldoConsolidadoxPedidos = ESNULO(RSAUX!SaldoConsolidadoxPedidos, 0)
 Else
   VGflagconversioncodigo = 0
End If
If clave < 4 Then
    If Buscar_en_BD Then
        VGUsuario = txtUserName
        VGPass = txtPassword
        'Ingresa Datos a la Tabla Menu
        Dim Clsmenu As New ClassMenu
        Set Clsmenu.conexion = VGConfig
        MDIPrincipal.StatusBar1.Panels.item(1) = "Empresa:  " & VGParametros.empresacodigo & " - " & UCase(VGParametros.NomEmpresa)
        MDIPrincipal.StatusBar1.Panels.item(2) = "Servidor :  " & UCase(VGParamSistem.Servidor)
        MDIPrincipal.StatusBar1.Panels.item(3) = "B.de Datos:  " & UCase(VGParamSistem.BDEmpresa)
        MDIPrincipal.StatusBar1.Panels.item(4) = "Usuario:  " & UCase(txtUserName.text)
        MDIPrincipal.StatusBar1.Panels.item(5) = "Fecha:  " & VGParamSistem.fechatrabajo
        MDIPrincipal.StatusBar1.Panels.item(6) = "Pto de Venta:  " & VGParametros.puntovta & " - " & RsPvta!puntovtadescripcion
               
        vgtipo = TIPOSISTEMA.INVENTARIOS
        Clsmenu.TablaMenu = "si_menu"
        Clsmenu.CrearTablaMenu
        Clsmenu.TabaMenuDet = "si_menuusuarios"
        Clsmenu.TablaMenu = "si_menu"
        Call Clsmenu.HabilitarMenuNom(VGUsuario)
        Unload Me
    Else
        MsgBox "La contraseña no es válida. Vuelva a intentarlo", vbInformation, "Inicio de sesión"
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
    
            reg.Open "Select * from Administrador", VGConfig, adOpenStatic
            
            If reg.RecordCount <> 0 Then
                reg.MoveFirst
                Do While Not reg.EOF
                    If UCase(Trim(txtUserName)) = UCase(reg.Fields("ADM_NOMBRE")) And UCase(RTrim$(txtPassword)) = DECODIFICA(reg.Fields("ADM_PASSWORD"), NUMMAGICO) Then
                            Buscar_en_BD = True
                            VGUsuario = UCase(txtUserName)
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
    reg.Open "Select * from USUARIO ", VGConfig, adOpenStatic
     Buscar_en_BD = False
     If reg.RecordCount <> 0 Then
     reg.MoveFirst
     Do While Not reg.EOF
        If UCase(RTrim$(txtUserName)) = UCase(reg.Fields("USU_CODIGO")) And UCase(RTrim$(txtPassword)) = DECODIFICA(reg.Fields("USU_PASSWORD"), NUMMAGICO) Then
            VGUsuario = UCase(txtUserName)
            VGPass = txtPassword
            Buscar_en_BD = True
            Exit Do
        End If
        reg.MoveNext
        If reg.EOF Then Exit Do
     Loop
    Else
        Buscar_en_BD = False
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
    rs.Open "Select * from EMPRESA WHERE  EMP_RAZON_NOMBRE = '" & Trim(Combo1.text) & "' and empresaflaginventarios= 1", VGConfig, adOpenStatic
    If Not rs.EOF Then
      VGCodEmpresa = rs("EMP_CODIGO")
      VGParametros.sistemamultiempresas = ESNULO(rs!multiempresas, False)
      VGParametros.RucEmpresa = ESNULO(REG1("EMP_RUC_DOCUMENTO"), "")
      VGCodigo = ESNULO(rs!codigoproducto, "")
      VGParametros.MontoexeneradoLiqCompra = numero(rs!MontoexoneradoLiqCompra)
      VGParametros.VGporcentajeimpto = numero(rs!Porcentajeimpuesto)
      VGmodovta = ESNULO(rs!modovtacodigo, "00")
      VGParametros.nombreguia = ESNULO(rs!nombreguiaremision, "al_guiaremision")
      VGParametros.nombrefactura = ESNULO(rs!nombrefactura, "vt_factura")
      VGParametros.multiboletas = ESNULO(rs!multiboletas, False)
      VGParametros.multifacturas = ESNULO(rs!multifacturas, False)
      VGParametros.multiguias = ESNULO(rs!multiguiasremision, False)
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
      VGCodEmpresa = REG1("EMP_CODIGO")
      VGParametros.NomEmpresa = REG1("EMP_RAZON_NOMBRE")
      VGParametros.RucEmpresa = REG1("EMP_RUC_DOCUMENTO")
    End If
    rs.Close
        
    Call adicionarcampos
    Call LlenarListaEmpresas
    Call LlenaPtoVta
    Set rs = VGCNx.Execute("select stockcomp from vt_parametroventa")
    stockcomp = ESNULO(rs.Fields(0), False)
    rs.Close
    RSQL = "Select  * From co_sistema"
    Set rs = New ADODB.Recordset
    Set rs = VGCNx.Execute(RSQL)
    If rs.RecordCount > 0 Then
        VGParametros.PermiteRequerimientos = ESNULO(rs("permiterequerimientos"), 0)
        VGParametros.PermiteIngresosconRequerimientos = ESNULO(rs("permiteIngresosconrequerimientos"), 0)
    End If
    
End If
Exit Sub

ERRAR:
MsgBox "Ocurrio un Error," & error & " debe Actualizar su Base de Datos e Ingrese Nuevamente al Sistema", vbInformation, "Aviso.."
Resume Next
End Sub

Sub LlenaPtoVta()
Dim rs As ADODB.Recordset

Set rs = VGCNx.Execute("select puntovtacodigo,puntovtadescripcion from vt_puntoventa ")
Combo4.Clear
Do While Not rs.EOF
    Combo4.AddItem rs(0) & " " & rs(1)
    rs.MoveNext
Loop

Combo4.ListIndex = 0
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub


Private Sub Combo3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub


Private Sub Combo4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub


Private Sub DTPfecha_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{tab}"
End Sub


Private Sub Form_Load()
Me.Cls
clave = 1
central Me
ADOCONECTAR
DTPfecha.Value = Date
SQL = " select * from si_sistema where tipodesistema =" & TIPOSISTEMA.INVENTARIOS & ""
Set REG1 = VGConfig.Execute(SQL)
If REG1.RecordCount > 0 Then Label1(0) = RTrim(REG1!tipodesistemadescripcion) + " : " + REG1!anno + "." + REG1!Version
Set REG1 = New ADODB.Recordset
REG1.Open "Select * from EMPRESA where empresaflaginventarios= 1 order by EMP_CODIGO ", VGConfig, adOpenStatic
VGParametros.sistemamultiempresas = REG1!multiempresas
Combo2.Visible = ESNULO(VGParametros.sistemamultiempresas, False)
LlenarListBox

cmdOK.Picture = MDIPrincipal.ImageList1.ListImages("Entrar").Picture
cmdCancel.Picture = MDIPrincipal.ImageList1.ListImages("Retornar").Picture

If Combo1.ListCount > 0 Then Combo1.ListIndex = 0
If Combo3.ListCount > 0 Then Combo3.ListIndex = 0

End Sub

Private Sub Form_Unload(Cancel As Integer)
'End
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
If REG1.EOF Then Exit Sub
If REG1.BOF Then Exit Sub
Do While Not REG1.EOF
   If Not IsNull(REG1.Fields("EMP_RAZON_NOMBRE")) Then
       Combo1.AddItem REG1.Fields("EMP_RAZON_NOMBRE")
   End If
     REG1.MoveNext
     If REG1.EOF Then Exit Do
Loop
REG1.MoveFirst
Combo1.ListIndex = 0
End Sub

