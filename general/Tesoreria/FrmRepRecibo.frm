VERSION 5.00
Object = "{4D137D9C-00A6-4458-9B46-7E95DB76D55B}#9.0#0"; "TextFer.ocx"
Begin VB.Form FrmRepRecibo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Imprimir Recibo Tesoreria"
   ClientHeight    =   660
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5910
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   660
   ScaleWidth      =   5910
   Begin TextFer.TxFer TxRecibo 
      Height          =   315
      Left            =   195
      TabIndex        =   0
      Top             =   135
      Width           =   2625
      _ExtentX        =   4630
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
      MaxLength       =   6
      Text            =   ""
      ColorIlumina    =   14546937
      Valor           =   ""
      NoCaracteres    =   "0123456789"
      SignodeMiles    =   -1  'True
      NoRangoCadena   =   -1  'True
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   330
      Left            =   4230
      TabIndex        =   2
      Top             =   150
      Width           =   1290
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   330
      Left            =   2895
      TabIndex        =   1
      Top             =   150
      Width           =   1290
   End
End
Attribute VB_Name = "FrmRepRecibo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdaceptar_Click()
  Call imprimir
End Sub

Private Sub CmdCancelar_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  Me.Width = 6000
  Me.Height = 1545
End Sub

Private Sub imprimir()
  Call ImprimirRecibo(TxRecibo.Text)
End Sub
