VERSION 5.00
Begin VB.Form FrmIngreso 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9705
   DrawStyle       =   6  'Inside Solid
   FillColor       =   &H000080FF&
   ForeColor       =   &H000000FF&
   Icon            =   "FrmIngreso.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   9705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   4095
      Left            =   5280
      TabIndex        =   0
      Top             =   120
      Width           =   4095
      Begin VB.CommandButton cAcepta 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Acepta"
         Height          =   1065
         Index           =   0
         Left            =   885
         MaskColor       =   &H00C0E0FF&
         Picture         =   "FrmIngreso.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2820
         Width           =   1095
      End
      Begin VB.CommandButton cCancela 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Cancela"
         Height          =   1065
         Index           =   1
         Left            =   2145
         MaskColor       =   &H00C0E0FF&
         Picture         =   "FrmIngreso.frx":15FE0
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   2820
         Width           =   1095
      End
      Begin VB.ComboBox Combo2 
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1290
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1395
         Width           =   1845
      End
      Begin VB.TextBox MText 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   0
         Left            =   1275
         MaxLength       =   8
         TabIndex        =   3
         Top             =   1890
         Width           =   1395
      End
      Begin VB.TextBox MText 
         Appearance      =   0  'Flat
         ForeColor       =   &H00C00000&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   1275
         MaxLength       =   8
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   2310
         Width           =   1395
      End
      Begin VB.ComboBox Combo1 
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1275
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   960
         Width           =   2610
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CUENTAS POR COBRAR"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   330
         Index           =   5
         Left            =   240
         TabIndex        =   11
         Top             =   240
         Width           =   3375
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   555
         Index           =   0
         Left            =   0
         Picture         =   "FrmIngreso.frx":28725
         Stretch         =   -1  'True
         Top             =   120
         Width           =   4515
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Index           =   0
         Left            =   150
         TabIndex        =   10
         Top             =   1950
         Width           =   735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Contraseña :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   9
         Top             =   2355
         Width           =   1185
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Grupo :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Index           =   2
         Left            =   165
         TabIndex        =   8
         Top             =   990
         Width           =   825
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Pto  Venta :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Index           =   3
         Left            =   150
         TabIndex        =   7
         Top             =   1470
         Width           =   1185
      End
   End
   Begin VB.Image Image1 
      Height          =   3570
      Index           =   1
      Left            =   360
      Picture         =   "FrmIngreso.frx":3801C
      Top             =   480
      Width           =   4530
   End
End
Attribute VB_Name = "FrmIngreso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adll As New dllgeneral.dll_general

Private Sub cAcepta_Click(Index As Integer)
Dim xsql As New ADODB.Recordset
   If Len(RTrim$(Combo1.Text)) = 0 Then
       MsgBox "Ingrese la Empresa", vbInformation, MsgTitle
       Combo1.SetFocus
       Exit Sub
   End If
   If Len(RTrim$(Combo2.Text)) = 0 Then
       MsgBox "Ingrese el Punto de Venta", vbInformation, MsgTitle
       Combo2.SetFocus
       Exit Sub
     Else
       VGparametros.puntovta = Left(Combo2.Text, 2)
   End If
   Set xsql = vgconfig.Execute("select * from usuario where USU_CODIGO='" & MText(0) & "'")
   If xsql.RecordCount = 0 Then
        MsgBox "Usuario No valido...Verifique!!!", vbInformation, MsgTitle
'FIXIT: Reemplazar la función 'UCase' con la función 'UCase$'.                             FixIT90210ae-R9757-R1B8ZE
   ElseIf UCase$(RTrim$(MText(1))) <> DECODIFICA(xsql.Fields("USU_PASSWORD"), NUMMAGICO) Then
        MsgBox "Contraseña errada .!!!", vbInformation, "Sistemas"
        Exit Sub
   End If
   g_Empresa = adll.ComboDato(Combo1.Text)
   g_DetalleEmpresa = Combo1.Text
   g_ptoventa = adll.ComboDato(Combo2.Text)
   g_usuario = RTrim$(MText(0))
   Call adicionarcampos
   Call Cargar_Parametros_Funcionales
   Dim Clsmenu As New ClassMenu
   Set Clsmenu.conexion = vgconfig
   VGtipo = cobrar
        Clsmenu.TablaMenu = "si_menu"
        Clsmenu.CrearTablaMenu
        Clsmenu.TabaMenuDet = "si_menuusuarios"
        Clsmenu.TablaMenu = "si_menu"
        Call Clsmenu.HabilitarMenuNom(g_usuario)
   If MDIPrincipal.Visible = False Then
      MDIPrincipal.Show
      Unload Me
   Else
      MostrarForm MDIPrincipal, "M"
      Call Cargar_Parametros_Funcionales
      Unload FrmIngreso
   End If
End Sub

Private Sub cCancela_Click(Index As Integer)
   If MDIPrincipal.Visible = True Then
        Unload Me
   Else
       End
   End If
End Sub

Private Sub Combo1_click()
Dim rs As ADODB.Recordset
If Combo1.ListCount > 0 Then
    Set rs = New ADODB.Recordset
    SQL = "Select * from EMPRESA WHERE  EMP_RAZON_NOMBRE = '" & RTrim$(Combo1.Text) & "' "
    SQL = SQL & " and empresaflagcobrar= 1"
    rs.Open (SQL), vgconfig, adOpenStatic
    If Not rs.EOF Then
      VGCODEMPRESA = rs("EMP_CODIGO")
      VGparametros.NomEmpresa = rs("EMP_RAZON_NOMBRE")
      VGparametros.RucEmpresa = Escadena(rs("EMP_RUC_DOCUMENTO"))
      If rs!empresaflagcobrar = 1 Then VGparamsistem.bdempresa = rs!empresabasecobrar
    End If
    rs.Close
    Set VGCNx = New ADODB.Connection
    VGCNx.CursorLocation = adUseClient
    VGCNx.CommandTimeout = 0
    VGCNx.ConnectionTimeout = 0
    VGCNx.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & VGparamsistem.Usuario & ";Password=" & VGparamsistem.pwd & ";Initial Catalog=" & VGparamsistem.bdempresa & ";Data Source=" & VGparamsistem.servidor
    VGCNx.Open
    If VGparamsistem.bdempresaCT <> VGparamsistem.bdempresa Then
       VGparamsistem.bdempresaCT = VGparamsistem.bdempresa
       Set VGcnxCT = New ADODB.Connection
       VGcnxCT.CursorLocation = adUseClient
       VGcnxCT.CommandTimeout = 0
       VGcnxCT.ConnectionTimeout = 0
       VGcnxCT.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & VGparamsistem.usuarioCT & ";Password=" & VGparamsistem.pwdCT & ";Initial Catalog=" & VGparamsistem.bdempresaCT & ";Data Source=" & VGparamsistem.servidorCT
       VGcnxCT.Open
    End If
End If
Call adll.llenacombo(Combo2, "select puntovtacodigo,puntovtadescripcion from vt_puntoventa", VGCNx)
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    Seguir Combo1, KeyAscii
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
    Seguir Combo2, KeyAscii
End Sub

Private Sub Form_Load()
    MostrarForm Me, "C1"
    Call CargaIngreso
End Sub

Public Sub CargaIngreso()
 Dim reg1 As New ADODB.Recordset
 Dim X As String
 reg1.Open "Select * from EMPRESA where empresaflagcobrar= 1 order by EMP_CODIGO ", vgconfig, adOpenStatic
 If reg1.EOF Then Exit Sub
 If reg1.BOF Then Exit Sub
 Do While Not reg1.EOF
    If Not IsNull(reg1.Fields("EMP_RAZON_NOMBRE")) Then
        Combo1.AddItem reg1.Fields("EMP_RAZON_NOMBRE")
    End If
    reg1.MoveNext
   If reg1.EOF Then Exit Do
 Loop
 reg1.MoveFirst
 Combo1.ListIndex = 0
End Sub


Private Sub MText_KeyPress(Index As Integer, KeyAscii As Integer)
  Seguir MText(Index), KeyAscii
End Sub

Private Sub MText_LostFocus(Index As Integer)
'FIXIT: Reemplazar la función 'UCase' con la función 'UCase$'.                             FixIT90210ae-R9757-R1B8ZE
   MText(Index) = UCase$(MText(Index))
End Sub
