VERSION 5.00
Object = "{4D137D9C-00A6-4458-9B46-7E95DB76D55B}#9.0#0"; "TextFer.ocx"
Begin VB.Form frmRepMovimientoCuentas 
   Caption         =   "Movimientos Cuentas"
   ClientHeight    =   2610
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5370
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2610
   ScaleWidth      =   5370
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   360
      Left            =   2828
      TabIndex        =   5
      Top             =   1860
      Width           =   1260
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   360
      Left            =   1283
      TabIndex        =   4
      Top             =   1860
      Width           =   1170
   End
   Begin TextFer.TxFer txtCuenta 
      Height          =   330
      Left            =   870
      TabIndex        =   1
      Top             =   465
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   582
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
      Text            =   ""
      Valor           =   ""
   End
   Begin TextFer.TxFer txtNivel 
      Height          =   315
      Left            =   870
      TabIndex        =   3
      Top             =   975
      Width           =   570
      _ExtentX        =   1005
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
      Text            =   ""
      Valor           =   ""
   End
   Begin VB.Label Label2 
      Caption         =   "Nivel"
      Height          =   270
      Left            =   165
      TabIndex        =   2
      Top             =   1065
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Cuenta:"
      Height          =   255
      Left            =   150
      TabIndex        =   0
      Top             =   540
      Width           =   615
   End
End
Attribute VB_Name = "frmRepMovimientoCuentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAceptar_Click()
  Call Impresion
End Sub

Private Sub cmdCancelar_Click()
   Unload Me
End Sub

Private Sub Form_Load()
  txtCuenta.Text = "(''94'',''95'',''97'')"
  txtNivel.Text = "2"
End Sub

Sub Impresion()
'FIXIT: Declare 'arrform' and 'arrparm' con un tipo de datos de enlace en tiempo de compilación     FixIT90210ae-R1672-R1B8ZE
Dim arrform(2) As Variant, arrparm() As Variant
    ReDim arrparm(5)
    arrparm(0) = VGParamSistem.BDEmpresa
    arrparm(1) = VGParamSistem.Anoproceso
    arrparm(2) = Trim$(txtCuenta.Text)
    arrparm(3) = Trim$(txtNivel.Text)
    arrform(0) = "@TituloReporte='" & "Resumen Movimientos" & "'"
    arrform(1) = "@Mes='" & "Enero - Diciembre" & " - " & VGParamSistem.Anoproceso & "'"
    Call ImpresionRptProc("RepResumenMovimientosCuentas.rpt", arrform, arrparm)
End Sub
