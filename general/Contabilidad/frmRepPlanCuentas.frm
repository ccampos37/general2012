VERSION 5.00
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Begin VB.Form frmRepPlanCuentas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Plan de Cuentas"
   ClientHeight    =   3972
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   5544
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3972
   ScaleWidth      =   5544
   Begin VB.CommandButton cmdBotones 
      Caption         =   "&Aceptar"
      Height          =   345
      Index           =   0
      Left            =   1568
      TabIndex        =   9
      Top             =   3465
      Width           =   1110
   End
   Begin VB.CommandButton cmdBotones 
      Caption         =   "&Cancelar"
      Height          =   345
      Index           =   1
      Left            =   2873
      TabIndex        =   8
      Top             =   3465
      Width           =   1110
   End
   Begin VB.Frame Frame3 
      Caption         =   "Cuenta Contable"
      Height          =   675
      Left            =   0
      TabIndex        =   7
      Top             =   2490
      Width           =   5550
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuda1 
         Height          =   330
         Left            =   60
         TabIndex        =   10
         Top             =   240
         Width           =   5400
         _ExtentX        =   9525
         _ExtentY        =   572
         XcodMaxLongitud =   0
         xcodwith        =   950
         NomTabla        =   "ct_cuenta"
         ListaCampos     =   "cuentacodigo(1),cuentadescripcion(1)"
         XcodCampo       =   "cuentacodigo"
         XListCampo      =   "cuentadescripcion"
         ListaCamposDescrip=   "Código,Descripción"
         ListaCamposText =   "cuentacodigo,cuentadescripcion"
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Parámetros de Impresión"
      Height          =   915
      Left            =   0
      TabIndex        =   3
      Top             =   1380
      Width           =   5550
      Begin VB.ComboBox cboNiveles 
         Height          =   315
         Left            =   2010
         TabIndex        =   6
         Top             =   240
         Width           =   3315
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Estructurado"
         Height          =   300
         Index           =   1
         Left            =   90
         TabIndex        =   5
         Top             =   510
         Width           =   2145
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Niveles"
         Height          =   300
         Index           =   0
         Left            =   105
         TabIndex        =   4
         Top             =   255
         Width           =   2145
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo de Impresión"
      Height          =   945
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   5535
      Begin VB.OptionButton Option1 
         Caption         =   "Plan de Cuentas Resumido"
         Height          =   195
         Index           =   1
         Left            =   105
         TabIndex        =   2
         Top             =   570
         Width           =   2310
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Plan de Cuentas Total"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   285
         Width           =   2310
      End
   End
End
Attribute VB_Name = "frmRepPlanCuentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
  Option1(0).Value = True
  Option2(0).Value = True
  Call Llenarcboniveles
  Call ConfiguraForm
End Sub

Sub Llenarcboniveles()
Dim i As Integer
  For i = 1 To VGnumnivelescuenta
    cboNiveles.AddItem "NIVEL " & Format(i, "0#")
 Next
End Sub

Private Sub Option2_Click(INDEX As Integer)
  Select Case INDEX
    Case 0: cboNiveles.Enabled = True
    Case 1: cboNiveles.Enabled = False
  End Select
End Sub

Private Sub CMDBOTONES_Click(INDEX As Integer)
  Select Case INDEX
    Case 0
    
    Case 1: Unload Me
  
  End Select

End Sub

Sub ConfiguraForm()
  Me.Width = 5670
  Me.Height = 4380
  'Me.Left = (MDIPrincipal.Width - Me.Width) / 2
  'Me.Top = (MDIPrincipal.Height - Me.Height) / 2
  Ctr_Ayuda1.conexion VGcnx
End Sub

