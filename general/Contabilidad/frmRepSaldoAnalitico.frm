VERSION 5.00
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "CONTROLAYUDA.OCX"
Begin VB.Form frmRepSaldoAnalitico 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Saldo Analítico"
   ClientHeight    =   4635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7140
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   7140
   Begin VB.CommandButton cmdBotones 
      Caption         =   "&Cancelar"
      Height          =   330
      Index           =   1
      Left            =   3788
      TabIndex        =   16
      Top             =   4095
      Width           =   1230
   End
   Begin VB.CommandButton cmdBotones 
      Caption         =   "&Aceptar"
      Height          =   330
      Index           =   0
      Left            =   2123
      TabIndex        =   10
      Top             =   4095
      Width           =   1230
   End
   Begin VB.Frame Frame1 
      Caption         =   "Seleccionar Opción"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   15
      TabIndex        =   0
      Top             =   390
      Width           =   7125
      Begin VB.OptionButton optOpcion 
         Caption         =   "Resumido"
         Height          =   225
         Index           =   1
         Left            =   135
         TabIndex        =   2
         Top             =   555
         Width           =   3000
      End
      Begin VB.OptionButton optOpcion 
         Caption         =   "Detallado"
         Height          =   225
         Index           =   0
         Left            =   135
         TabIndex        =   1
         Top             =   285
         Width           =   3000
      End
   End
   Begin VB.Frame fraDetallado 
      Caption         =   "Detallado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2430
      Left            =   15
      TabIndex        =   3
      Top             =   1395
      Width           =   7125
      Begin VB.CheckBox Check2 
         Caption         =   "Todos los Códigos Analíticos"
         Height          =   210
         Left            =   75
         TabIndex        =   12
         Top             =   1080
         Width           =   2610
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuda1 
         Height          =   330
         Left            =   1695
         TabIndex        =   19
         Top             =   525
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   582
         XcodMaxLongitud =   0
         xcodwith        =   800
         NomTabla        =   "ct_cuenta"
         ListaCampos     =   "cuentacodigo(1),cuentadescripcion(1)"
         XcodCampo       =   "cuentacodigo"
         XListCampo      =   "cuentadescripcion"
         ListaCamposDescrip=   "Código,Descripción"
         ListaCamposText =   "cuentacodigo,cuentadescripcion"
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuda2 
         Height          =   330
         Left            =   3510
         TabIndex        =   17
         Top             =   1395
         Width           =   3570
         _ExtentX        =   6297
         _ExtentY        =   582
         XcodMaxLongitud =   0
         xcodwith        =   700
      End
      Begin VB.Frame Frame5 
         Height          =   60
         Left            =   0
         TabIndex        =   5
         Top             =   915
         Width           =   7095
      End
      Begin VB.Frame Frame4 
         Height          =   120
         Left            =   30
         TabIndex        =   8
         Top             =   1755
         Width           =   7065
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Todas las Cuentas Contables"
         Height          =   405
         Left            =   90
         TabIndex        =   15
         Top             =   180
         Width           =   2610
      End
      Begin VB.CheckBox chkPendiente 
         Alignment       =   1  'Right Justify
         Caption         =   "Solamente Pendientes"
         Height          =   390
         Left            =   60
         TabIndex        =   9
         Top             =   1920
         Width           =   1440
      End
      Begin VB.ComboBox cboTipoAnalitico 
         Height          =   315
         Left            =   1095
         TabIndex        =   7
         Top             =   1395
         Width           =   1800
      End
      Begin VB.Label Label4 
         Caption         =   "Código"
         Height          =   270
         Left            =   2955
         TabIndex        =   18
         Top             =   1455
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Tipo Analítico"
         Height          =   240
         Left            =   60
         TabIndex        =   6
         Top             =   1470
         Width           =   1005
      End
      Begin VB.Label Label1 
         Caption         =   "Cuenta Contable"
         Height          =   255
         Left            =   75
         TabIndex        =   4
         Top             =   600
         Width           =   1335
      End
   End
   Begin VB.Frame fraresumido 
      Caption         =   "Resumido"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   990
      Left            =   15
      TabIndex        =   11
      Top             =   1395
      Width           =   7125
      Begin VB.CheckBox chkAmbitoCuenta 
         Caption         =   "Todas las Cuentas Contables"
         Height          =   240
         Left            =   75
         TabIndex        =   14
         Top             =   240
         Width           =   2640
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuda3 
         Height          =   330
         Left            =   1575
         TabIndex        =   20
         Top             =   510
         Width           =   5205
         _ExtentX        =   9181
         _ExtentY        =   582
         XcodMaxLongitud =   0
         xcodwith        =   800
         NomTabla        =   "ct_cuenta"
         ListaCampos     =   "cuentacodigo(1),cuentadescripcion(1)"
         XcodCampo       =   "cuentacodigo"
         XListCampo      =   "cuentadescripcion"
         ListaCamposDescrip=   "Código,Descripción"
         ListaCamposText =   "cuentacodigo,cuentadescripcion"
      End
      Begin VB.Label Label3 
         Caption         =   "Cuenta Contable"
         Height          =   255
         Left            =   75
         TabIndex        =   13
         Top             =   585
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmRepSaldoAnalitico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const FLAG_RESUMIDO As Boolean = True
Const FLAG_DETALLE As Boolean = False

Private Sub Form_Load()
  Call LlenarcboTipoAnalitico
  Call ConfiguraForm(True)
End Sub

Sub LlenarcboTipoAnalitico()
 Dim dllgen As New dllgeneral.dll_general
 Dim rs As ADODB.Recordset
 
 Set rs = VGcnx.Execute("Select tipoanaliticocodigo,tipoanaliticodescripcion from ct_tipoanalitico")
 cboTipoAnalitico.Clear
 While Not rs.EOF
   cboTipoAnalitico.AddItem rs(1)
   rs.MoveNext
 Wend
 Set dllgen = New dllgeneral.dll_general
End Sub

Private Sub Check1_Click()
 If Check1.Value = 0 Then
   Label1.Visible = True
   Ctr_Ayuda1.Visible = True
 Else
   Label1.Visible = False
   Ctr_Ayuda1.Visible = False
 End If
End Sub

Private Sub Check2_Click()
 If Check2.Value = 0 Then
   Ctr_Ayuda2.Visible = True
   Label4.Visible = True
 Else
   Ctr_Ayuda2.Visible = False
   Label4.Visible = False
 End If
End Sub

Private Sub chkAmbitoCuenta_Click()
  Select Case chkAmbitoCuenta.Value
    Case 0:
      Ctr_Ayuda2.Visible = True
      Label4.Visible = True
      
    Case 1:
      Ctr_Ayuda2.Visible = False
      Label4.Visible = False
  End Select
  
End Sub

Private Sub optOpcion_Click(Index As Integer)
  Select Case Index
    Case 0:
      fraDetallado.Visible = True
      fraresumido.Visible = False
      Call ConfiguraForm(FLAG_DETALLE)
      
    Case 1:
      fraDetallado.Visible = False
      fraresumido.Visible = True
      Call ConfiguraForm(FLAG_RESUMIDO)
  
  End Select
End Sub

Sub ConfiguraForm(flag As Boolean)
 'Flag .T.=Modo Resumido  .F.=Modo Detallado
 Me.Width = 7260
 Me.Height = IIf(flag = True, 3840, 5070)
 cmdBotones(0).Top = IIf(flag, 2760, 4080)
 cmdBotones(1).Top = IIf(flag, 2760, 4080)
 
 Check1.Value = 1
 Check2.Value = 1
 
 Ctr_Ayuda1.conexion VGcnx
 Ctr_Ayuda2.conexion VGcnx
 
End Sub

Private Sub cmdBotones_Click(Index As Integer)
 Select Case Index
  Case 0:
  
  Case 1: Unload Me
 
 End Select

End Sub

