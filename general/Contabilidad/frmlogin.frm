VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{4D137D9C-00A6-4458-9B46-7E95DB76D55B}#9.0#0"; "TextFer.ocx"
Begin VB.Form frmlogin 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   Caption         =   "Formulario de Ingreso"
   ClientHeight    =   4680
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8730
   ControlBox      =   0   'False
   Icon            =   "frmlogin.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4680
   ScaleWidth      =   8730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   4455
      Index           =   1
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   8535
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H000000FF&
         Height          =   3870
         Index           =   0
         Left            =   4440
         TabIndex        =   14
         Top             =   360
         Width           =   3930
         Begin VB.CommandButton CmdAceptar 
            BackColor       =   &H00C0FFFF&
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
            Height          =   1065
            Left            =   600
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   2520
            Width           =   1215
         End
         Begin VB.CommandButton CmdCancelar 
            BackColor       =   &H00C0FFFF&
            Caption         =   "&Salir"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1065
            Left            =   2220
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   2520
            Width           =   1215
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   1110
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   240
            Width           =   2655
         End
         Begin VB.ComboBox Combo2 
            Height          =   315
            Left            =   1110
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   750
            Visible         =   0   'False
            Width           =   2655
         End
         Begin MSComCtl2.DTPicker DTPfecha 
            Height          =   285
            Left            =   1215
            TabIndex        =   5
            Top             =   1965
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   503
            _Version        =   393216
            Format          =   97648641
            CurrentDate     =   37508
         End
         Begin TextFer.TxFer TxUser 
            Height          =   315
            Left            =   1215
            TabIndex        =   3
            Top             =   1185
            Width           =   1845
            _ExtentX        =   3254
            _ExtentY        =   556
            Object.CausesValidation=   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxLength       =   8
            Text            =   ""
            Valor           =   ""
         End
         Begin TextFer.TxFer TxPwd 
            Height          =   315
            Left            =   1215
            TabIndex        =   4
            Top             =   1545
            Width           =   1845
            _ExtentX        =   3254
            _ExtentY        =   556
            Object.CausesValidation=   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxLength       =   8
            PasswordChar    =   "*"
            Text            =   ""
            Valor           =   ""
         End
         Begin VB.Label Lbgrupo 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFFF&
            Caption         =   "Grupo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   150
            TabIndex        =   19
            Top             =   315
            Width           =   525
         End
         Begin VB.Label Lbempresa 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFFF&
            Caption         =   "Empresa :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   18
            Top             =   825
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Usuario :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   255
            TabIndex        =   17
            Top             =   1260
            Width           =   780
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Password :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   255
            TabIndex        =   16
            Top             =   1650
            Width           =   945
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha de Trabajo :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   255
            TabIndex        =   15
            Top             =   1875
            Width           =   870
         End
      End
      Begin VB.Image Image1 
         Height          =   3750
         Left            =   240
         Picture         =   "frmlogin.frx":0442
         Top             =   360
         Width           =   6285
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   1170
      Left            =   195
      TabIndex        =   10
      Top             =   1125
      Width           =   3900
   End
   Begin VB.Frame framaño 
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      Height          =   870
      Left            =   90
      TabIndex        =   0
      Top             =   6525
      Visible         =   0   'False
      Width           =   3885
      Begin VB.CommandButton cmdGenerar 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         Caption         =   "&Generar Año"
         Height          =   330
         Left            =   1095
         MaskColor       =   &H000000FF&
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   435
         UseMaskColor    =   -1  'True
         Width           =   1785
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "El año seleccionado no esta generado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   285
         TabIndex        =   9
         Top             =   165
         Width           =   3555
      End
   End
   Begin MSComctlLib.ImageList ImgList2 
      Left            =   2610
      Top             =   5715
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmlogin.frx":A227
            Key             =   "Grabar"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmlogin.frx":197E7
            Key             =   "Imprimir"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmlogin.frx":2846B
            Key             =   "Modificar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmlogin.frx":378C4
            Key             =   "Borrar"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmlogin.frx":4ABD8
            Key             =   "Cancelar"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmlogin.frx":5AD22
            Key             =   "Retornar"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmlogin.frx":6D477
            Key             =   "Camara"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmlogin.frx":813D4
            Key             =   "Crear"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmlogin.frx":9981C
            Key             =   "Celular"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmlogin.frx":A8E85
            Key             =   "Ver"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmlogin.frx":C67A3
            Key             =   "Arbitrios"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmlogin.frx":D69B9
            Key             =   "Autoav"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmlogin.frx":DB7BB
            Key             =   "Facturar"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmlogin.frx":EB025
            Key             =   "Entrar"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmlogin.frx":101009
            Key             =   "Convenio"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmlogin.frx":10C023
            Key             =   "Generar"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Iniciar sesion"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   270
      Left            =   135
      TabIndex        =   12
      Top             =   585
      Width           =   1365
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sistema de Contabilidad"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   360
      Left            =   90
      TabIndex        =   11
      Top             =   180
      Width           =   3660
   End
End
Attribute VB_Name = "frmlogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim VGvardllgen As New dllgeneral.dll_general

Private Sub cmdAceptar_Click()
Dim CLMENU As ClassMenu
Dim NUSER As Integer
Dim tccambio As Double
    Call CargarParametros
    If Not Validaraño Then
        framaño.Visible = True
        framaño.Top = 4050
        CmdAceptar.Top = 5000
        CmdCancelar.Top = 5000
        frmlogin.Height = 6500
        Exit Sub
       Else
        framaño.Visible = False
        framaño.Top = 6400
    End If
    If Not VERIFICAUSUARIO Then Exit Sub
    tccambio = XRecuperaTipoCambio(Format(DTPfecha, "dd/mm/yyyy"), Venta, VGCNx)
    If tccambio = 0 Then
        MsgBox "No existe tipo de cambio para esta fecha", vbInformation
    End If
    MDIPrincipal.StatusBar1.Panels(1).Text = "Mes Proceso : " & VGvardllgen.DesMes(Month(DTPfecha))
    MDIPrincipal.StatusBar1.Panels(2).Text = "Año Proceso :" & Year(DTPfecha)
    MDIPrincipal.StatusBar1.Panels(3).Text = "Fecha de Trabajo (" & VGParamSistem.FechaTrabajo & ")"
    MDIPrincipal.StatusBar1.Panels(4).Text = "Tipo Cambio  (" & Format(tccambio, "#.000") & ")"
    VGUsuario = UCase(TxUser.Text)
        Dim Clsmenu As New ClassMenu
        Set Clsmenu.Conexion = VGConfig
        VGtipo = contab
        Clsmenu.TablaMenu = "si_menu"
        Clsmenu.CrearTablaMenu
        Clsmenu.TabaMenuDet = "si_menuusuarios"
        Clsmenu.TablaMenu = "si_menu"
        Call Clsmenu.HabilitarMenuNom(VGUsuario)
       Call CargarParametrosContabilidad
       Call adicionarcamposCT
    If TxUser.Text <> "" Then
        VGParamSistem.Usuario = TxUser.Text
        VGParamSistem.Pwd = ""
        MDIPrincipal.Caption = MDIPrincipal.Caption & " Usuario : " & TxUser.Text
     Else
        MDIPrincipal.Caption = MDIPrincipal.Caption & " Usuario : Sin Usuario "
    End If
   Unload Me
End Sub

Private Sub CargarParametros()
Dim rssql As New ADODB.Recordset
    VGParamSistem.Anoproceso = Format(Year(DTPfecha), "0000")
    VGParamSistem.Mesproceso = Format(Month(DTPfecha), "00")
    VGParamSistem.TablaCabcomprob = "ct_cabcomprob" & Year(DTPfecha)
    VGParamSistem.TablaDetcomprob = "ct_detcomprob" & Year(DTPfecha)
 ''   VGParamSistem.RutaReport = App.Path
    VGParamSistem.FechaTrabajo = DTPfecha.Value
    Set rssql = VGCNx.Execute("select * from co_multiempresas where empresacodigo='" & VGParametros.empresacodigo & "'")
    If rssql.RecordCount > 0 Then VGParametros.RucEmpresa = ESNULO(rssql!empresaruc, "")
End Sub
Private Function Validaraño() As Boolean
Validaraño = True
Dim rsaux As ADODB.Recordset
Dim cab As Boolean, det As Boolean, msg As String
    Set rsaux = New ADODB.Recordset
    'Verificar que existan cabecera de comprobante
    rsaux.Open "select name from sysobjects where name in ('" & _
                        VGParamSistem.TablaCabcomprob & "','" & _
                        VGParamSistem.TablaDetcomprob & "')", VGCNx
    If rsaux.RecordCount <= 1 Then
        MsgBox "No estan aperturadas la tablas para este año", vbExclamation
        Validaraño = False
    End If
End Function

Private Sub cmdCancelar_Click()
    VGSalir = True
    Unload Me
End Sub

Private Sub cmdGenerar_Click()
    frmannos.Visible = False
    frmannos.DTPanno.Value = DTPfecha.Value
    frmannos.cmdGenerar_Click
    framaño.Visible = False
    Unload frmannos
End Sub
Private Sub Combo1_Click()
Dim RSQL As String
Dim rs As ADODB.Recordset
On Local Error GoTo ERRAR
If Combo1.ListCount > 0 Then
    Set rs = New ADODB.Recordset
    SQL = "Select * from EMPRESA WHERE  EMP_RAZON_NOMBRE = '" & Trim$(Combo1.Text) & "' "
    SQL = SQL & " and empresaflagcontabilidad= 1"
    Set rs = VGConfig.Execute(SQL)
    Lbgrupo.Caption = ""
    If Not rs.EOF Then
      VGParametros.NomEmpresa = rs("EMP_RAZON_NOMBRE")
      VGParametros.RucEmpresa = ESNULO(rs("EMP_RUC_DOCUMENTO"), "")
      If ESNULO(rs!multiempresas, False) = True Then Lbgrupo.Caption = "Grupo"
      VGParametros.multiempresas = ESNULO(rs!multiempresas, False)
      VGCodEmpresa = rs("EMP_CODIGO")
      If Trim$(rs!empresabasecontabilidad) <> "" Then
         VGParamSistem.BDEmpresa = rs!empresabasecontabilidad
         Set VGCNx = New ADODB.Connection
         VGCNx.CursorLocation = adUseClient
         VGCNx.CommandTimeout = 0
         VGCNx.ConnectionTimeout = 0
         VGCNx.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & VGParamSistem.Usuario & ";Password=" & VGParamSistem.Pwd & ";Initial Catalog=" & VGParamSistem.BDEmpresa & ";Data Source=" & VGParamSistem.Servidor
         VGCNx.Open
      End If
    End If
    If Lbgrupo.Caption = "Grupo" Then
       Lbempresa.Visible = True
       Combo2.Visible = True
     Else
       Lbgrupo.Caption = "Empresa"
       Lbempresa.Visible = False
       Combo2.Visible = False
    End If
    Call LlenarListaempresas
    rs.Close
 
    Call CargarParametrosContabilidad
End If
Exit Sub

ERRAR:
MsgBox "Ocurrio un Error," & error & " debe Actualizar su Base de Datos e Ingrese Nuevamente al Sistema", vbInformation, "Aviso.."
Resume Next
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Public Sub Combo2_Click()
      VGParametros.empresacodigo = Left(Combo2.Text, 2)
      VGParametros.NomEmpresa = Right(Combo2.Text, Len(Combo2.Text) - 2)
      MDIPrincipal.Caption = "Sistema de Contabilidad - " & VGParametros.NomEmpresa
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub


Private Sub DTPfecha_KeyDown(KeyCode As Integer, Shift As Integer)
' If KeyCode = 13 Then SendKeys "{tab}"
End Sub


Private Sub Form_Load()
DTPfecha.Value = Date
Call LlenarListBox
CmdAceptar.Picture = ImgList2.ListImages.Item("Entrar").Picture
CmdCancelar.Picture = ImgList2.ListImages.Item("Retornar").Picture
'
'framaño.Top = 6400
'CmdAceptar.Top = 4050
'cmdCancelar.Top = 4050
'frmlogin.Height = 5600
'
'framaño.Visible = True
'framaño.Top = 4050
'CmdAceptar.Top = 5000
'cmdCancelar.Top = 5000
'frmlogin.Height = 6500

End Sub
Private Function VERIFICAUSUARIO() As Boolean
    Dim RSPASS As New ADODB.Recordset
    Dim Pwd As String
    Dim CLMENU As ClassMenu
    Set CLMENU = New ClassMenu
    CLMENU.TablaUsu = "SI_USUARIO"
    
    'cuando no existe usuarios
    VERIFICAUSUARIO = False
    Set RSPASS = New ADODB.Recordset
    RSPASS.Open "SELECT * FROM " & UCase$(CLMENU.TablaUsu), VGConfig, adOpenKeyset
    If RSPASS.RecordCount = 0 Then
        VERIFICAUSUARIO = True
        Exit Function
    End If
    'VALIDANDO SI EXISTE EL USUARIO
    Set RSPASS = New ADODB.Recordset
    RSPASS.Open "SELECT * FROM " & UCase$(CLMENU.TablaUsu) & " WHERE USUARIOCODIGO='" & TxUser.Text & "'", VGConfig, adOpenKeyset
    If RSPASS.RecordCount = 0 Then
        MsgBox "NO SE ENCUENTRA EL USUARIO ", vbExclamation
        TxUser.SetFocus
        Exit Function
    End If
    
    'VALIDANDO SI EXISTE EL PWD
    Pwd = CODIFICA(TxPwd.Text, 5)
    Set RSPASS = Nothing
    SQL = "SELECT * FROM " & UCase$(CLMENU.TablaUsu) & " WHERE USUARIOCODIGO='" & TxUser.Text & "'"
    SQL = SQL & " AND USUarioPASSWORD='" & Pwd & "'"
    Set RSPASS = VGConfig.Execute(SQL)
    If RSPASS.RecordCount = 0 Then
        MsgBox "LA CONTRASEÑA ES INCORRECTA", vbExclamation
        TxPwd.SetFocus
        Exit Function
    End If
    VERIFICAUSUARIO = True
End Function

Private Sub Form_Unload(Cancel As Integer)
   'End
End Sub

Public Sub LlenarListBox()
Dim REG1 As New ADODB.Recordset
Set REG1 = VGConfig.Execute("Select * from EMPRESA where empresaflagcontabilidad= 1 order by EMP_CODIGO ")
If REG1.BOF Then Exit Sub
Do While Not REG1.EOF
   If Not IsNull(REG1.Fields("EMP_RAZON_NOMBRE")) Then
      Combo1.AddItem REG1.Fields("EMP_RAZON_NOMBRE")
   End If
   REG1.MoveNext
Loop
If Combo1.ListCount > 0 Then Combo1.ListIndex = 0
End Sub

Public Sub LlenarListaempresas()
Dim REG1 As New ADODB.Recordset
Set REG1 = New ADODB.Recordset
Dim multiempresa As Integer
Set REG1 = VGCNx.Execute("Select * from co_multiempresas where empresacodigo<>'00'  ")
Combo2.Clear
If REG1.EOF Then Exit Sub
If REG1.BOF Then Exit Sub
Do While Not REG1.EOF
   Combo2.AddItem REG1.Fields("empresacodigo") + " " + REG1.Fields("empresadescripcion")
   REG1.MoveNext
Loop

REG1.MoveFirst
Combo2.ListIndex = 0


End Sub




Private Sub TxPwd_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub


Private Sub TxUser_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   TxUser.Text = UCase$(TxUser.Text)
   SendKeys "{tab}"
End If
End Sub


