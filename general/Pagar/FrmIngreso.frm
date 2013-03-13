VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmIngreso 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5730
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   10185
   ControlBox      =   0   'False
   DrawStyle       =   6  'Inside Solid
   FillColor       =   &H00800000&
   Icon            =   "FrmIngreso.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   10185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox MText 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   6840
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox MText 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Index           =   0
      Left            =   6840
      MaxLength       =   8
      TabIndex        =   0
      Top             =   2400
      Width           =   1815
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Left            =   6840
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1800
      Width           =   3015
   End
   Begin VB.CommandButton cCancela 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Cancela"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Index           =   1
      Left            =   7440
      MaskColor       =   &H00C0E0FF&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4320
      Width           =   1365
   End
   Begin VB.CommandButton cAcepta 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Acepta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Index           =   0
      Left            =   5670
      MaskColor       =   &H00C0E0FF&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4320
      Width           =   1365
   End
   Begin MSComCtl2.DTPicker DTPfecha 
      Height          =   495
      Left            =   6840
      TabIndex        =   9
      Top             =   3600
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   108789761
      CurrentDate     =   37508
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha Trabajo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   540
      Index           =   4
      Left            =   5400
      TabIndex        =   10
      Top             =   3600
      Width           =   1035
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   5280
      Left            =   360
      Picture         =   "FrmIngreso.frx":0442
      Stretch         =   -1  'True
      Top             =   150
      Width           =   4305
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SISTEMA DE CUENTAS POR PAGAR"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   915
      Left            =   5160
      TabIndex        =   8
      Top             =   165
      Width           =   3915
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Empresa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   2
      Left            =   5400
      TabIndex        =   7
      Top             =   1860
      Width           =   1110
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Contraseña"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   1
      Left            =   5370
      TabIndex        =   6
      Top             =   3120
      Width           =   1470
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   0
      Left            =   5400
      TabIndex        =   5
      Top             =   2460
      Width           =   1020
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00800000&
      BorderWidth     =   2
      FillColor       =   &H00C0FFFF&
      FillStyle       =   0  'Solid
      Height          =   3915
      Index           =   0
      Left            =   5040
      Shape           =   4  'Rounded Rectangle
      Top             =   1470
      Width           =   4875
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0FFFF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   4065
      Index           =   1
      Left            =   4950
      Shape           =   4  'Rounded Rectangle
      Top             =   1380
      Width           =   5025
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
If Len(Trim$(Combo1.Text)) = 0 Then
       MsgBox "Ingrese la Empresa", vbInformation, MsgTitle
       Combo1.SetFocus
       Exit Sub
   End If
   g_Empresa = adll.ComboDato(Combo1.Text)
   g_DetalleEmpresa = Combo1.Text
   VGusuario = Trim$(MText(0))
   Set xsql = VGconfig.Execute("select * from usuario where USU_CODIGO='" & MText(0) & "'")
   If xsql.RecordCount = 0 Then
        MsgBox "Usuario No valido...Verifique!!!", vbInformation, MsgTitle
   ElseIf UCase$(RTrim$(MText(1))) <> DECODIFICA(xsql.Fields("USU_PASSWORD"), NUMMAGICO) Then
        MsgBox "Contraseña errada .!!!", vbInformation, "Sistemas"
        Exit Sub
   End If
   
   Call adicionarcampos
   Call Cargar_Parametros_Funcionales
   Dim Clsmenu As New ClassMenu
   Set Clsmenu.conexion = VGconfig
   VGtipo = pagar
        Clsmenu.TablaMenu = "si_menu"
        Clsmenu.CrearTablaMenu
        Clsmenu.TabaMenuDet = "si_menuusuarios"
        Clsmenu.TablaMenu = "si_menu"
        Call Clsmenu.HabilitarMenuNom(VGusuario)
   VGparamsistem.Mesproceso = Format(Month(DTPFecha), "00")
   VGparamsistem.Anoproceso = Format(Year(DTPFecha), "0000")
    
  If MDIPrincipal.Visible = False Then
      MDIPrincipal.Show
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
    SQL = "Select * from EMPRESA WHERE  EMP_RAZON_NOMBRE = '" & Trim$(Combo1.Text) & "' "
    SQL = SQL & " and empresaflagpagar= 1"
    rs.Open (SQL), VGconfig, adOpenStatic
    If Not rs.EOF Then
      VGCODEMPRESA = rs("EMP_CODIGO")
      VGparametros.NomEmpresa = rs("EMP_RAZON_NOMBRE")
      VGparametros.RucEmpresa = ESNULO(rs("EMP_RUC_DOCUMENTO"), "")
      If rs!empresaflagpagar = 1 Then VGparamsistem.BDEmpresa = rs!empresabasepagar
    End If
    rs.Close
    Set VGCNx = New ADODB.Connection
    VGCNx.CursorLocation = adUseClient
    VGCNx.CommandTimeout = 0
    VGCNx.ConnectionTimeout = 0
    VGCNx.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & VGparamsistem.Usuario & ";Password=" & VGparamsistem.PWD & ";Initial Catalog=" & VGparamsistem.BDEmpresa & ";Data Source=" & VGparamsistem.Servidor
    VGCNx.Open
    If VGparamsistem.BDEmpresaCT <> VGparamsistem.BDEmpresa Then
       VGparamsistem.BDEmpresaCT = VGparamsistem.BDEmpresa
       Set VGcnxCT = New ADODB.Connection
       VGcnxCT.CursorLocation = adUseClient
       VGcnxCT.CommandTimeout = 0
       VGcnxCT.ConnectionTimeout = 0
       VGcnxCT.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & VGparamsistem.UsuarioCT & ";Password=" & VGparamsistem.PwdCT & ";Initial Catalog=" & VGparamsistem.BDEmpresaCT & ";Data Source=" & VGparamsistem.ServidorCT
       VGcnxCT.Open
    End If
End If
End Sub

Private Sub Form_Load()
    MostrarForm Me, "C1"
    Call CargaIngreso
End Sub

Public Sub CargaIngreso()

 Dim reg1 As New ADODB.Recordset
 reg1.Open "Select * from EMPRESA where empresaflagpagar= 1 order by EMP_CODIGO ", VGconfig, adOpenStatic
 If reg1.EOF Then
    reg1.Close
    reg1.Open "Select top 1 * from EMPRESA order by EMP_CODIGO ", VGconfig, adOpenStatic
    Combo1.AddItem reg1.Fields("EMP_RAZON_NOMBRE")
  Else
    Do While Not reg1.EOF
       If Not IsNull(reg1.Fields("EMP_RAZON_NOMBRE")) Then
          Combo1.AddItem reg1.Fields("EMP_RAZON_NOMBRE")
       End If
       reg1.MoveNext
    Loop
    reg1.MoveFirst
 End If
 Combo1.ListIndex = 0
   DTPFecha.Value = Format(Now, "dd/mm/yyyy")
   VGparamsistem.FechaTrabajo = DTPFecha.Value

 
End Sub

Private Sub MText_KeyPress(Index As Integer, KeyAscii As Integer)

  Seguir MText(Index), KeyAscii
End Sub

Private Sub MText_LostFocus(Index As Integer)
   MText(Index) = UCase$(MText(Index))
End Sub
