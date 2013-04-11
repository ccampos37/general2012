VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{4D137D9C-00A6-4458-9B46-7E95DB76D55B}#9.0#0"; "TextFer.ocx"
Begin VB.Form frmlogin 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SISTEMA DE COMPRAS "
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8655
   Icon            =   "frmlogin.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   8655
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frameempresa 
      BackColor       =   &H00C0FFFF&
      Height          =   975
      Left            =   4680
      TabIndex        =   14
      Top             =   1200
      Width           =   3855
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label LbGrupo 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Grupo  / Empresa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   16
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame framaño 
      BackColor       =   &H00C0FFFF&
      Height          =   1215
      Left            =   4680
      TabIndex        =   11
      Top             =   3840
      Visible         =   0   'False
      Width           =   3855
      Begin VB.CommandButton cmdGenerar 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         Caption         =   "&Generar Año"
         Height          =   330
         Left            =   1095
         MaskColor       =   &H000000FF&
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   675
         UseMaskColor    =   -1  'True
         Width           =   1785
      End
      Begin VB.Label Label5 
         BackColor       =   &H0080FF80&
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
         Height          =   285
         Left            =   210
         TabIndex        =   13
         Top             =   285
         Width           =   3555
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   330
      Left            =   6885
      TabIndex        =   5
      Top             =   5100
      Width           =   1110
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   330
      Left            =   5535
      TabIndex        =   4
      Top             =   5100
      Width           =   1110
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      Height          =   1095
      Left            =   4680
      TabIndex        =   7
      Top             =   120
      Width           =   3855
      Begin VB.Label Label2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Ingrese el Usuario, Password y Elija la fecha de trabajo."
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
         Height          =   585
         Left            =   1485
         TabIndex        =   8
         Top             =   225
         Width           =   1905
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   135
         Picture         =   "frmlogin.frx":0442
         Top             =   390
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   645
         Picture         =   "frmlogin.frx":0884
         Top             =   240
         Width           =   480
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H000000FF&
      Height          =   1815
      Left            =   4680
      TabIndex        =   0
      Top             =   2040
      Width           =   3855
      Begin MSComCtl2.DTPicker DTPfecha 
         Height          =   285
         Left            =   1095
         TabIndex        =   3
         Top             =   1125
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   503
         _Version        =   393216
         Format          =   109051905
         CurrentDate     =   37508
      End
      Begin TextFer.TxFer TxUser 
         Height          =   315
         Left            =   1095
         TabIndex        =   1
         Top             =   240
         Width           =   2580
         _ExtentX        =   4551
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
         SaltarAlEnter   =   -1  'True
         Valor           =   ""
      End
      Begin TextFer.TxFer TxPwd 
         Height          =   315
         Left            =   1095
         TabIndex        =   2
         Top             =   585
         Width           =   2580
         _ExtentX        =   4551
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
         PasswordChar    =   "*"
         Text            =   ""
         SaltarAlEnter   =   -1  'True
         Valor           =   ""
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0FFFF&
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
         Left            =   135
         TabIndex        =   10
         Top             =   1035
         Width           =   870
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
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
         Left            =   135
         TabIndex        =   9
         Top             =   690
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
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
         Left            =   135
         TabIndex        =   6
         Top             =   300
         Width           =   780
      End
   End
   Begin VB.Image Image1 
      Height          =   5040
      Index           =   1
      Left            =   -360
      Picture         =   "frmlogin.frx":0CC6
      Top             =   240
      Width           =   5235
   End
End
Attribute VB_Name = "frmlogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAceptar_Click()
Dim tccambio As Double
VGUsuario = Trim(TxUser.Text)
VGfecha = DTPfecha.Value
    
Dim CLMENU As ClassMenu
Call CargarParametros
Call adicionarcampos
Call CargarParametrosCompras
'   If Not Validaraño Then
'        framaño.Visible = True
'        Exit Sub
'       Else
'        framaño.Visible = False
'    End If
   tccambio = XRecuperaTipoCambio(DTPfecha.Value, Venta, VGCnxCT)
    If tccambio = 0 Then
        If Not VGParametros.permite_tc Then
            MsgBox "No existe Tipo de cambio para esta fecha ", vbInformation
          Else
            MsgBox "No existe Tipo de cambio para esta fecha " & Chr(13) & _
                   "Por lo tanto no podra ingresar al sistema y se Cerrara", vbInformation
            End
        End If
    End If
    Call adicionacamposCO
    MDIPrincipal.StatusBar1.Panels(3).Text = "Fecha de Trabajo (" & VGParamSistem.FechaTrabajo & ")"
    MDIPrincipal.StatusBar1.Panels(4).Text = "Tipo Cambio  (" & Format(tccambio, "#.000") & ")"
    
    Set CLMENU = New ClassMenu
    If Not VERIFICAUSUARIO Then Exit Sub
    If TxUser.Text <> "" Then
        VGParamSistem.Usuario = TxUser.Text
        VGParamSistem.PWD = ""
        VGtipo = compras
        CLMENU.TablaMenu = "si_menu"
        CLMENU.CrearTablaMenu
        CLMENU.TabaMenuDet = "si_menuusuarios"
        CLMENU.TablaMenu = "si_menu"
        Call CLMENU.HabilitarMenuNom(Trim(TxUser.Text))
        MDIPrincipal.Caption = MDIPrincipal.Caption & " Usuario : " & TxUser.Text
       Else
        MDIPrincipal.Caption = MDIPrincipal.Caption & " Usuario : Sin Usuario "
    End If
    Unload Me
    
    
End Sub


Private Function VERIFICAUSUARIO() As Boolean
    Dim RSPASS As New ADODB.Recordset
    Dim PWD As String
    Dim CLMENU As ClassMenu
    Set CLMENU = New ClassMenu
    CLMENU.TablaUsu = "USUARIO"
      
    'cuando no existe usuarios
    VERIFICAUSUARIO = False
   'VALIDANDO SI EXISTE EL USUARIO
    Set RSPASS = New ADODB.Recordset
    RSPASS.Open "SELECT * FROM " & UCase(CLMENU.TablaUsu), VGConfig
    If RSPASS.RecordCount = 0 Then
       VERIFICAUSUARIO = True
       Exit Function
    End If
    Set RSPASS = New ADODB.Recordset
    RSPASS.Open "SELECT * FROM " & UCase(CLMENU.TablaUsu) & " WHERE USU_CODIGO='" & TxUser.Text & "'", VGConfig
    If RSPASS.RecordCount = 0 Then
        MsgBox "NO SE ENCUENTRA EL USUARIO ", vbExclamation
        TxUser.SetFocus
        Exit Function
    End If
    
    'VALIDANDO SI EXISTE EL PWD
    
    PWD = CODIFICA(TxPwd.Text, 5)
    Set RSPASS = Nothing
    RSPASS.Open "SELECT * FROM " & UCase(CLMENU.TablaUsu) & " WHERE USU_CODIGO='" & TxUser.Text & _
    "' AND USU_PASSWORD='" & PWD & "'", VGConfig, adOpenKeyset
    If RSPASS.RecordCount = 0 Then
        MsgBox "LA CONTRASEÑA ES INCORRECTA", vbExclamation
        TxPwd.SetFocus
        Exit Function
    End If
    VERIFICAUSUARIO = True
End Function

Private Sub CargarParametros()
    VGParamSistem.Anoproceso = Format(Year(DTPfecha), "0000")
    VGParamSistem.Mesproceso = Format(Month(DTPfecha), "00")
    VGParamSistem.TablaCabcomprob = "co_cabeceraprovisiones"
    VGParamSistem.TablaDetcomprob = "co_detalleprovisiones"
    VGParamSistem.FechaTrabajo = DTPfecha.Value
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

Private Sub CmdCancelar_Click()
    End
End Sub

Private Sub cmdGenerar_Click()
    frmannos.Visible = False
    
    frmannos.DTPanno.Value = DTPfecha.Value
    frmannos.cmdGenerar_Click
    framaño.Visible = False
    Unload frmannos
End Sub

Public Sub Combo1_click()
Dim rs As ADODB.Recordset
If Combo1.ListCount > 0 Then
    Set rs = New ADODB.Recordset
    SQL = "Select * from EMPRESA WHERE  EMP_RAZON_NOMBRE = '" & Trim(Combo1.Text) & "' "
    SQL = SQL & " and empresaflagcompras= 1"
    LbGrupo.Caption = "Empresa"
    rs.Open (SQL), VGConfig, adOpenStatic
    If Not rs.EOF Then
      VGCodEmpresa = rs("EMP_CODIGO")
      VGParametros.NomEmpresa = rs("EMP_RAZON_NOMBRE")
      If rs!empresaflagcompras = 1 Then VGParamSistem.BDEmpresa = Trim(rs!empresabasecompras)
      If ESNULO(rs!multiempresas, False) = True Then LbGrupo.Caption = "Grupo"
      VGtipolicencia = rs!tipodelicencias
      VGfechalicencia = rs!fechalimitelicencias
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
       Set VGCnxCT = New ADODB.Connection
       VGCnxCT.CursorLocation = adUseClient
       VGCnxCT.CommandTimeout = 0
       VGCnxCT.ConnectionTimeout = 0
       VGCnxCT.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & VGParamSistem.UsuarioCT & ";Password=" & VGParamSistem.PwdCT & ";Initial Catalog=" & VGParamSistem.BDEmpresaCT & ";Data Source=" & VGParamSistem.ServidorCT
       VGCnxCT.Open
    End If

End If
End Sub


Private Sub Form_Load()
   DTPfecha.Value = Date
   LlenarListBox
End Sub
Public Sub LlenarListBox()
Dim REG1 As New ADODB.Recordset
Set REG1 = New ADODB.Recordset
Dim multiempresa As Integer
LbGrupo.Caption = "Empresa"
REG1.Open "Select * from EMPRESA where empresaflagcompras= 1 order by EMP_CODIGO ", VGConfig, adOpenStatic
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


