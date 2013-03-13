VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmIngreso 
   Appearance      =   0  'Flat
   BackColor       =   &H80000002&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5265
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   3780
   ControlBox      =   0   'False
   DrawStyle       =   6  'Inside Solid
   FillColor       =   &H00FFFFFF&
   Icon            =   "FrmIngreso.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   3780
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox Combo2 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1395
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1935
      Width           =   1845
   End
   Begin VB.TextBox MText 
      ForeColor       =   &H00C00000&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   1410
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   2760
      Width           =   855
   End
   Begin VB.TextBox MText 
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   0
      Left            =   1410
      MaxLength       =   8
      TabIndex        =   2
      Top             =   2355
      Width           =   855
   End
   Begin VB.ComboBox Combo1 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1410
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1485
      Width           =   1845
   End
   Begin VB.CommandButton cCancela 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Salir"
      Height          =   1095
      Index           =   1
      Left            =   1980
      MaskColor       =   &H00C0E0FF&
      Picture         =   "FrmIngreso.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3810
      Width           =   1185
   End
   Begin VB.CommandButton cAcepta 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Aceptar"
      Height          =   1095
      Index           =   0
      Left            =   420
      MaskColor       =   &H000000FF&
      Picture         =   "FrmIngreso.frx":3256
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3810
      Width           =   1185
   End
   Begin MSComCtl2.DTPicker DTPfecha 
      Height          =   285
      Left            =   1380
      TabIndex        =   4
      Top             =   3150
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   503
      _Version        =   393216
      Format          =   38207489
      CurrentDate     =   37508
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SISTEMA DE TESORERIA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   600
      TabIndex        =   10
      Top             =   360
      Width           =   2475
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha Trabajo"
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
      Height          =   420
      Index           =   4
      Left            =   270
      TabIndex        =   12
      Top             =   3150
      Width           =   1035
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Oficina"
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
      Height          =   255
      Index           =   3
      Left            =   270
      TabIndex        =   11
      Top             =   2010
      Width           =   1035
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Empresa"
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
      Height          =   255
      Index           =   2
      Left            =   270
      TabIndex        =   9
      Top             =   1560
      Width           =   1035
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Contraseña"
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
      Height          =   255
      Index           =   1
      Left            =   270
      TabIndex        =   8
      Top             =   2805
      Width           =   1035
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario"
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
      Height          =   255
      Index           =   0
      Left            =   270
      TabIndex        =   6
      Top             =   2385
      Width           =   1035
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      Height          =   1275
      Left            =   -30
      Top             =   -30
      Width           =   3855
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
Dim tccambio As Integer
   If Len(Trim(Combo1.Text)) = 0 Then
       MsgBox "Ingrese la Empresa", vbInformation, MsgTitle
       Combo1.SetFocus
       Exit Sub
   End If
   If Len(Trim(Combo2.Text)) = 0 Then
       MsgBox "Ingrese la Oficina", vbInformation, MsgTitle
       Combo2.SetFocus
       Exit Sub
   End If
'-----------------------------------------------------------------------------------
   If adll.VerificaDatoExistente(VGConfig, "select * from si_usuario") > 0 Then
      If adll.VerificaDatoExistente(VGConfig, "select * from si_usuario where usuariocodigo='" & MText(0) & "'") = 0 Then
         MsgBox "No existe usuario....Verifique!!!", vbInformation, MsgTitle
         Call adll.Enfoquetexto(MText(0))
         Exit Sub
       ElseIf adll.VerificaDatoExistente(VGConfig, "select * from si_usuario where usuariocodigo='" & MText(0) & "' and usuarioPassword='" & CODIFICA(MText(1), NUMMAGICO) & "' ") = 0 Then
             MsgBox "Contraseña Errada No valido...Verifique!!!", vbInformation, MsgTitle
             Call adll.Enfoquetexto(MText(1))
             Exit Sub
       End If
   End If
'-----------------------------------------------------------------------------------

   VGCodEmpresa = adll.ComboDato(Combo1.Text)
   VGUsuario = Trim(MText(0))
   VGParamSistem.AnoProceso = Format(Year(DTPfecha), "0")
   VGParamSistem.MesProceso = Format(Month(DTPfecha), "0")
   VGParamSistem.TablaCabcomprob = "co_cabeceraprovisiones"
   VGParamSistem.tabladetcomprob = "co_detalleprovisiones"
   VGParamSistem.fechatrabajo = DTPfecha.Value
   
   tccambio = XRecuperaTipoCambio(Format(DTPfecha, "dd/mm/yyyy"), Venta, VGcnxCT)
   
   If tccambio = 0 Then
        MsgBox "No existe tipo de cambio para esta fecha", vbInformation
        Exit Sub
    End If
    MDIPrincipal.Panel.Panels(4).Text = "Tipo Cambio  (" & Format(tccambio, "#.000") & ")"
    MDIPrincipal.Panel.Panels(3).Text = "Fecha de Trabajo (" & VGParamSistem.fechatrabajo & ")"
    
   Call adicionarcampos
   Call Cargar_Parametros_Funcionales
   
   Dim Clsmenu As New ClassMenu
   VGtipo = caja
   Clsmenu.TablaMenu = "si_menu"
   Clsmenu.CrearTablaMenu
   Clsmenu.TabaMenuDet = "si_menuusuarios"
   Clsmenu.TablaMenu = "si_menu"
   Call Clsmenu.HabilitarMenuNom(VGUsuario)
  
 
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

Private Sub Combo1_Click()
Dim rs As ADODB.Recordset
If Combo1.ListCount > 0 Then
    Set rs = New ADODB.Recordset
    SQL = "Select * from EMPRESA WHERE  EMP_RAZON_NOMBRE = '" & Trim(Combo1.Text) & "' "
    SQL = SQL & " and empresaflagtesoreria= 1"
    rs.Open (SQL), VGConfig, adOpenStatic
    If Not rs.EOF Then
'      VGCODEMPRESA = rs("EMP_CODIGO")
      If rs!empresaflagtesoreria = 1 Then VGParametros.NomEmpresa = rs("EMP_RAZON_NOMBRE")
      If rs!empresaflagtesoreria = 1 Then VGParametros.RucEmpresa = ESNULO(rs("EMP_RUC_Documento"), "")
      If rs!empresaflagtesoreria = 1 Then VGParamSistem.BDEmpresa = rs!empresabasetesoreria
 '     VGtipolicencia = rs!tipodelicencias
 '     VGfrechalicencia = rs!fechalimitelicencias
    End If
    rs.Close
    Set VGCNx = New ADODB.Connection
    VGCNx.CursorLocation = adUseClient
    VGCNx.CommandTimeout = 0
    VGCNx.ConnectionTimeout = 0
    VGCNx.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & VGParamSistem.Usuario & ";Password=" & VGParamSistem.Pwd & ";Initial Catalog=" & VGParamSistem.BDEmpresa & ";Data Source=" & VGParamSistem.Servidor
    VGCNx.Open
   Call adll.llenacombo(Combo2, "select vendedorcodigo,vendedornombres from cp_oficina", VGCNx)
     If VGParamSistem.BDEmpresaCT <> VGParamSistem.BDEmpresa Then
       VGParamSistem.BDEmpresaCT = VGParamSistem.BDEmpresa
       Set VGcnxCT = New ADODB.Connection
       VGcnxCT.CursorLocation = adUseClient
       VGcnxCT.CommandTimeout = 0
       VGcnxCT.ConnectionTimeout = 0
       VGcnxCT.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & VGParamSistem.UsuarioCT & ";Password=" & VGParamSistem.PwdCT & ";Initial Catalog=" & VGParamSistem.BDEmpresaCT & ";Data Source=" & VGParamSistem.ServidorCT
       VGcnxCT.Open
    End If
    
End If
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
    Seguir Combo2, KeyAscii
End Sub

Private Sub Combo2_LostFocus()
VGoficina = adll.ComboDato(Combo2.Text)
End Sub

Private Sub DTPfecha_KeyDown(KeyCode As Integer, Shift As Integer)
    Seguir DTPfecha, KeyCode
 End Sub

Private Sub DTPfecha_LostFocus()
   VGParamSistem.fechatrabajo = DTPfecha.Value
End Sub

Private Sub Form_Load()
    MostrarForm Me, "C1"
    Call CargaIngreso
End Sub

Public Sub CargaIngreso()
Dim reg1 As New ADODB.Recordset
 reg1.Open "Select * from EMPRESA where empresaflagtesoreria= 1 order by EMP_CODIGO ", VGConfig, adOpenStatic
 If reg1.EOF Then
    reg1.Close
    reg1.Open "Select top 1 * from EMPRESA order by EMP_CODIGO ", VGConfig, adOpenStatic
    Combo1.AddItem reg1.Fields("EMP_RAZON_NOMBRE")
    VGParametros.descripcion = reg1.Fields("emp_razon_nombre")
  Else
    Do While Not reg1.EOF
       If Not IsNull(reg1.Fields("EMP_RAZON_NOMBRE")) Then
          Combo1.AddItem reg1.Fields("EMP_RAZON_NOMBRE")
          VGParametros.descripcion = reg1.Fields("emp_razon_nombre")
       End If
       reg1.MoveNext
    Loop
    reg1.MoveFirst
 End If
 Combo1.ListIndex = 0
   Call adll.llenacombo(Combo2, "select vendedorcodigo,vendedornombres from cp_oficina", VGCNx)
   DTPfecha.Value = Format(Now, "dd/mm/yyyy")
   VGParamSistem.fechatrabajo = DTPfecha.Value
End Sub

Private Sub MText_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = 13 Then
    If Index = 0 And adll.VerificaDatoExistente(VGCNx, "select * from gr_usuario where usuariocodigo='" & MText(0) & "'") = 0 Then
       MsgBox "No existe usuario....Verifique!!!", vbInformation, MsgTitle
       Call adll.Enfoquetexto(MText(0))
       Exit Sub
    ElseIf Index = 1 And adll.VerificaDatoExistente(VGCNx, "select * from gr_usuario where usuariocodigo='" & MText(0) & "' and usuariopassword='" & MText(1) & "' ") = 0 Then
       MsgBox "Contraseña Errada y/o Usuario No valido...Verifique!!!", vbInformation, MsgTitle
       Call adll.Enfoquetexto(MText(1))
       Exit Sub

    End If
  End If
  Seguir MText(Index), KeyAscii
End Sub

Private Sub MText_LostFocus(Index As Integer)
   MText(Index) = UCase(MText(Index))
End Sub
