VERSION 5.00
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmIngreso 
   BackColor       =   &H0080FFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Jck -  Nextel 41*156*5229 / RPM *6906374 / RPC 993900810"
   ClientHeight    =   4965
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9135
   DrawStyle       =   6  'Inside Solid
   DrawWidth       =   2
   FillColor       =   &H00FFFFC0&
   ForeColor       =   &H008080FF&
   Icon            =   "FrmIngreso.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   9135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Height          =   4695
      Left            =   4920
      TabIndex        =   0
      Top             =   120
      Width           =   4095
      Begin VB.ComboBox Combo1 
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1470
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   630
         Width           =   2340
      End
      Begin VB.ComboBox Combo2 
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1470
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1350
         Width           =   2340
      End
      Begin VB.TextBox MText 
         Appearance      =   0  'Flat
         ForeColor       =   &H00C00000&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   1470
         MaxLength       =   8
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   2100
         Width           =   1410
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   1470
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   975
         Visible         =   0   'False
         Width           =   2340
      End
      Begin VB.TextBox MText 
         Appearance      =   0  'Flat
         ForeColor       =   &H00C00000&
         Height          =   285
         HideSelection   =   0   'False
         IMEMode         =   3  'DISABLE
         Index           =   0
         Left            =   1470
         MaxLength       =   8
         TabIndex        =   4
         Top             =   1740
         Width           =   1410
      End
      Begin VB.CommandButton CmdCancelar 
         BackColor       =   &H00FFFFFF&
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
         Height          =   1095
         Left            =   2640
         Picture         =   "FrmIngreso.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   3000
         Width           =   1215
      End
      Begin VB.CommandButton CmdAceptar 
         BackColor       =   &H00FFFFFF&
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
         Height          =   1095
         Left            =   840
         Picture         =   "FrmIngreso.frx":12A4F
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   3000
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker DTPfecha 
         Height          =   285
         Left            =   1470
         TabIndex        =   6
         Top             =   2520
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   503
         _Version        =   393216
         Format          =   100532225
         CurrentDate     =   37508
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyuCajero 
         Height          =   315
         Left            =   1470
         TabIndex        =   9
         Top             =   4515
         Visible         =   0   'False
         Width           =   2730
         _ExtentX        =   4815
         _ExtentY        =   556
         XcodMaxLongitud =   11
         xcodwith        =   200
         NomTabla        =   "vt_cajeros"
         TituloAyuda     =   "Busqueda de Codigos de cajeros"
         ListaCampos     =   "cajerocodigo(1),cajeroapellidos(1)"
         XcodCampo       =   "cajerocodigo"
         XListCampo      =   "cajeroapellidos"
         ListaCamposDescrip=   "Codigo,Apelliods"
         ListaCamposText =   "cajerocodigo,cajeroapellidos"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
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
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   0
         Left            =   270
         TabIndex        =   18
         Top             =   1755
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         Caption         =   "Contraseña :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   1
         Left            =   270
         TabIndex        =   17
         Top             =   2145
         Width           =   1095
      End
      Begin VB.Label lbgrupo 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         Caption         =   "Grupo :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   270
         TabIndex        =   16
         Top             =   660
         Width           =   645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         Caption         =   "Pto.  Venta :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   3
         Left            =   270
         TabIndex        =   15
         Top             =   1395
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SISTEMA DE FACTURACION"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   285
         Index           =   4
         Left            =   405
         TabIndex        =   14
         Top             =   240
         Width           =   3345
      End
      Begin VB.Label Lbempresa 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
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
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   270
         TabIndex        =   13
         Top             =   1020
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         Caption         =   "Fec. Trabajo :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   270
         TabIndex        =   12
         Top             =   2490
         Width           =   1215
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         Caption         =   "Cajero :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   270
         TabIndex        =   11
         Top             =   4560
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Version 2002.11"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   2
         Left            =   240
         TabIndex        =   10
         Top             =   4200
         Width           =   1290
      End
   End
   Begin VB.Image Image1 
      Height          =   4740
      Left            =   -3240
      Picture         =   "FrmIngreso.frx":28A23
      Top             =   120
      Width           =   8250
   End
End
Attribute VB_Name = "FrmIngreso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adll As New dllgeneral.dll_general
Dim tccambio As Double
Dim SQL As String


Private Sub cmdAceptar_Click()
Dim tccambio As Double
Dim xsql As New ADODB.Recordset
  tccambio = XRecuperaTipoCambio(DTPfecha.Value, Venta, VGcnxCT)
    If tccambio = 1 Then
       MsgBox "No existe Tipo de cambio para esta fecha " & Chr(13) & _
             "Por lo tanto no podra ingresar al sistema y se Cerrara", vbInformation
       Exit Sub
    End If
    VGParametros.cajerocodigo = "01"  ' Ctr_AyuCajero.xclave
   Ctr_AyuCajero.xclave = "01"
   If Len(Trim(Combo1.Text)) = 0 Then
       MsgBox "Ingrese la Empresa", vbInformation, MsgTitle
       Combo1.SetFocus
       Exit Sub
   End If
   If Len(Trim(Combo2.Text)) = 0 Then
       MsgBox "Ingrese el Punto de Venta", vbInformation, MsgTitle
       Combo2.SetFocus
       Exit Sub
   End If
   If adll.VerificaDatoExistente(VGConfig, "select * from si_usuario where usuariocodigo='" & MText(0) & "'") = 0 Then
       MsgBox "No existe usuario....Verifique!!!", vbInformation, MsgTitle
       Call adll.Enfoquetexto(MText(0))
       Exit Sub
   End If
   
   Set xsql = VGConfig.Execute("select * from si_usuario where usuariocodigo='" & MText(0) & "'")
   If xsql.RecordCount = 0 Then
        MsgBox "Usuario No valido...Verifique!!!", vbInformation, MsgTitle
   ElseIf UCase(RTrim$(MText(1))) <> DECODIFICA(xsql.Fields("USUARIOPASSWORD"), NUMMAGICO) Then
        MsgBox "Contraseña errada .!!!", vbInformation, "Sistemas"
        Exit Sub
   End If
   
   If Len(Trim(Ctr_AyuCajero.xclave)) = 0 Then
        MsgBox "Falta seleccionar Cajero", vbCritical, "Sistema"
        Ctr_AyuCajero.SetFocus
        Exit Sub
   End If
   
   If VGParametros.sistemamultiempresas Then
     VGParametros.empresacodigo = Left(Combo3.Text, 2)
   Else
     VGParametros.empresacodigo = "01"
   End If
   VGParametros.nomempresa = Right(RTrim(Combo3.Text), Len((RTrim(Combo3.Text))) - 2)
   g_ptoventa = adll.ComboDato(Combo2.Text)
   g_usuario = Trim(MText(0))
   VGParamSistem.AnoProceso = Format(Year(DTPfecha), "0000")
   VGParametros.mesproceso = Format(Month(DTPfecha), "00")
   Call Cargar_Parametros_Funcionales
  Call adicionarcampos
   
   Dim Clsmenu As New ClassMenu
   vgtipo = facturacion
        Clsmenu.TablaMenu = "si_menu"
        Clsmenu.CrearTablaMenu

        Clsmenu.TabaMenuDet = "si_menuusuarios"
        Clsmenu.TablaMenu = "si_menu"
        Call Clsmenu.HabilitarMenuNom(g_usuario)
  
   
   If MDIPrincipal.Visible = False Then
      MDIPrincipal.Show
   Else
      MostrarFormVentas MDIPrincipal, "M"
      Call Cargar_Parametros_Funcionales
      Unload FrmIngreso
   End If
MDIPrincipal.StatusBar1.Panels(3).Text = "  Fecha de Trabajo (" & VGParamSistem.FechaTrabajo & ")  "
MDIPrincipal.StatusBar1.Panels(4).Text = "  Tipo Cambio  (" & Format(tccambio, "#.000") & ")  "
MDIPrincipal.StatusBar1.Panels(5).Text = "   Base de datos  (" & VGCNx.DefaultDatabase & ")   "

End Sub

Private Sub cmdCancelar_Click()
'   If MDIPrincipal.Visible = True Then
  '      Unload Me
 '  Else
       End
 '  End If
End Sub



Private Sub Combo1_KeyPress(KeyAscii As Integer)
Seguir Combo1, KeyAscii
End Sub
Private Sub Combo1_Click()
Dim rs As ADODB.Recordset
If Combo1.ListCount > 0 Then
    Set rs = New ADODB.Recordset
    SQL = "Select * from EMPRESA WHERE  EMP_RAZON_NOMBRE = '" & Trim(Combo1.Text) & "' "
    SQL = SQL & " and empresaflagventas= 1"
    rs.Open (SQL), VGConfig, adOpenStatic
    If Not rs.EOF Then
      VGCodEmpresa = rs("EMP_CODIGO")
      VGParametros.sistemamultiempresas = ESNULO(rs!multiempresas, False)
      VGParametros.nomempresa = rs("EMP_RAZON_NOMBRE")
      VGParametros.RucEmpresa = ESNULO(rs("EMP_RUC_DOCUMENTO"), "")
      If rs!empresaflagventas = 1 Then VGParamSistem.BDEmpresa = rs!empresabaseventas
      VGParametros.multiboletas = ESNULO(rs!multiboletas, False)
      VGParametros.multifacturas = ESNULO(rs!multifacturas, False)
      VGParametros.multiguias = ESNULO(rs!multiguiasremision, False)
   End If
    rs.Close
    Set VGCNx = New ADODB.Connection
    VGCNx.CursorLocation = adUseClient
    VGCNx.CommandTimeout = 0
    VGCNx.ConnectionTimeout = 0
    VGCNx.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & VGParamSistem.Usuario & ";Password=" & VGParamSistem.PWD & ";Initial Catalog=" & VGParamSistem.BDEmpresa & ";Data Source=" & VGParamSistem.Servidor
    VGCNx.Open
    If VGParamSistem.BDEmpresaCT <> VGParamSistem.BDEmpresa Then
       VGParamSistem.BDEmpresaCT = VGParamSistem.BDEmpresa
       Set VGcnxCT = New ADODB.Connection
       VGcnxCT.CursorLocation = adUseClient
       VGcnxCT.CommandTimeout = 0
       VGcnxCT.ConnectionTimeout = 0
       VGcnxCT.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & VGParamSistem.UsuarioCT & ";Password=" & VGParamSistem.PwdCT & ";Initial Catalog=" & VGParamSistem.BDEmpresaCT & ";Data Source=" & VGParamSistem.ServidorCT
       VGcnxCT.Open
    End If
    If lbgrupo.Caption = "Grupo" Then
       Lbempresa.Visible = True
       Combo3.Visible = True
     Else
       lbgrupo.Caption = "Empresa"
       Lbempresa.Visible = False
       Combo3.Visible = False
    End If
    Combo3.Visible = VGParametros.sistemamultiempresas
    Lbempresa.Visible = VGParametros.sistemamultiempresas

    Call LlenarListaempresas
    Call adll.llenacombo(Combo2, "select puntovtacodigo,puntovtadescripcion from vt_puntoventa", VGCNx)
End If
End Sub
Private Sub Combo2_KeyPress(KeyAscii As Integer)
    Seguir Combo2, KeyAscii
End Sub

Private Sub Combo3_Click()
VGParametros.empresacodigo = Left(Combo3.Text, 2)
  Call adicionarcampos
End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
Seguir Combo3, KeyAscii
End Sub


Private Sub DTPfecha_Change()
   VGParamSistem.FechaTrabajo = DTPfecha.Value
   tccambio = XRecuperaTipoCambio(VGParamSistem.FechaTrabajo, Venta, VGCNx)
End Sub

Private Sub DTPfecha_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{tab}"
End Sub


Private Sub Form_Load()
    MostrarFormVentas Me, "C1"
    DTPfecha.Value = Date
    VGParamSistem.FechaTrabajo = DTPfecha.Value
    Call Ctr_AyuCajero.conexion(VGCNx)
    Call CargaIngreso
End Sub


Public Sub CargaIngreso()
   
 Dim REG1 As New ADODB.Recordset
 REG1.Open "Select * from EMPRESA where empresaflagventas= 1 order by EMP_CODIGO ", VGConfig, adOpenStatic
 VGParametros.sistemamultiempresas = REG1!multiempresas
 Combo3.Visible = ESNULO(VGParametros.sistemamultiempresas, False)
 Lbempresa.Visible = ESNULO(VGParametros.sistemamultiempresas, False)
 VGParametros.multiguias = ESNULO(REG1!multiguiasremision, False)
 VGParametros.multifacturas = ESNULO(REG1!multifacturas, False)
 VGParametros.multiboletas = ESNULO(REG1!multiboletas, False)
 If REG1.EOF Then Exit Sub
 If REG1.BOF Then Exit Sub
 REG1.MoveFirst
 Do While Not REG1.EOF
    If Not IsNull(REG1.Fields("EMP_RAZON_NOMBRE")) Then
        Combo1.AddItem REG1.Fields("EMP_RAZON_NOMBRE")
    End If
    REG1.MoveNext
   If REG1.EOF Then Exit Do
 Loop
 Combo1.ListIndex = 0
 Call adll.llenacombo(Combo2, "select puntovtacodigo,puntovtadescripcion from vt_puntoventa", VGCNx)

End Sub


Private Sub MText_KeyPress(Index As Integer, KeyAscii As Integer)
  Seguir MText(Index), KeyAscii
End Sub

Private Sub MText_LostFocus(Index As Integer)
   MText(Index) = UCase(MText(Index))
End Sub
Public Sub LlenarListaempresas()
Dim REG1 As New ADODB.Recordset
Set REG1 = New ADODB.Recordset
Dim multiempresa As Integer
Set REG1 = VGCNx.Execute("Select * from co_multiempresas where empresacodigo<>'00'  ")
Combo3.Clear
If REG1.EOF Then Exit Sub
If REG1.BOF Then Exit Sub
Do While Not REG1.EOF
   Combo3.AddItem REG1.Fields("empresacodigo") + " " + REG1.Fields("empresadescripcion")
   REG1.MoveNext
Loop

REG1.MoveFirst
Combo3.ListIndex = 0


End Sub

