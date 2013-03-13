VERSION 5.00
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Begin VB.Form frmRepPlanCuentasSaldosIniciales 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Saldos Iniciales"
   ClientHeight    =   3975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5640
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   5640
   Begin VB.Frame Frame1 
      Caption         =   "Tipo Impresión"
      Height          =   885
      Left            =   0
      TabIndex        =   8
      Top             =   255
      Width           =   5640
      Begin VB.OptionButton Option1 
         Caption         =   "Plan de Cuentas Total"
         Height          =   225
         Index           =   0
         Left            =   135
         TabIndex        =   10
         Top             =   270
         Width           =   2415
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Plan de Cuentas Resumido"
         Height          =   225
         Index           =   1
         Left            =   120
         TabIndex        =   9
         Top             =   540
         Width           =   2415
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Cuenta Contable"
      Height          =   705
      Left            =   0
      TabIndex        =   6
      Top             =   2550
      Width           =   5640
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuda1 
         Height          =   330
         Left            =   60
         TabIndex        =   7
         Top             =   225
         Width           =   5520
         _ExtentX        =   9737
         _ExtentY        =   582
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
   Begin VB.Frame Frame4 
      Caption         =   "Parámetros de Impresión"
      Height          =   915
      Left            =   0
      TabIndex        =   2
      Top             =   1470
      Width           =   5640
      Begin VB.ComboBox cboNiveles 
         Height          =   315
         Left            =   2340
         TabIndex        =   5
         Top             =   270
         Width           =   3240
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Niveles"
         Height          =   240
         Index           =   0
         Left            =   105
         TabIndex        =   4
         Top             =   300
         Width           =   1590
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Estructurado"
         Height          =   240
         Index           =   1
         Left            =   90
         TabIndex        =   3
         Top             =   555
         Width           =   1590
      End
   End
   Begin VB.CommandButton cmdBotones 
      Caption         =   "&Aceptar"
      Height          =   330
      Index           =   0
      Left            =   1635
      TabIndex        =   1
      Top             =   3480
      Width           =   1065
   End
   Begin VB.CommandButton cmdBotones 
      Caption         =   "&Cancelar"
      Height          =   330
      Index           =   1
      Left            =   2940
      TabIndex        =   0
      Top             =   3480
      Width           =   1065
   End
End
Attribute VB_Name = "frmRepPlanCuentasSaldosIniciales"
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

Private Sub Option2_Click(Index As Integer)
  Select Case Index
    Case 0: cboNiveles.Enabled = True
    Case 1: cboNiveles.Enabled = False
  End Select
End Sub

Private Sub cmdBotones_Click(Index As Integer)
  Select Case Index
    Case 0:
    Call imprimir
    Case 1: Unload Me
  
  End Select

End Sub

Sub ConfiguraForm()
  Me.Width = 5760
  Me.Height = 4380
  'Call CentrarForm(MDIPrincipal, Me)
  Ctr_Ayuda1.conexion VGCNx
End Sub
Sub imprimir()
'FIXIT: Declare 'arrparam' con un tipo de datos de enlace en tiempo de compilación         FixIT90210ae-R1672-R1B8ZE
Dim arrparam(3) As Variant
'FIXIT: Declare 'arrform' con un tipo de datos de enlace en tiempo de compilación          FixIT90210ae-R1672-R1B8ZE
Dim arrform(1) As Variant
    arrparam(0) = VGCNx.DefaultDatabase
    arrparam(1) = VGParametros.empresacodigo
    arrparam(2) = VGParamSistem.Anoproceso

arrform(0) = "ano='" & VGParamSistem.Anoproceso & "'"
    
Call ImpresionRptProc("ct_saldosiniciales.rpt", arrform, arrparam, , "Saldos iniciales")
End Sub
