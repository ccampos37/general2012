VERSION 5.00
Object = "{4D137D9C-00A6-4458-9B46-7E95DB76D55B}#9.0#0"; "textfer.ocx"
Begin VB.Form frmRepTransferencia 
   Caption         =   "Impresión de Transferencias"
   ClientHeight    =   675
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5610
   LinkTopic       =   "Form1"
   ScaleHeight     =   675
   ScaleWidth      =   5610
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   330
      Left            =   2813
      TabIndex        =   1
      Top             =   165
      Width           =   1290
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   330
      Left            =   4148
      TabIndex        =   2
      Top             =   165
      Width           =   1290
   End
   Begin TextFer.TxFer TxtTransf 
      Height          =   315
      Left            =   113
      TabIndex        =   0
      Top             =   150
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
      SaltarAlEnter   =   -1  'True
      Valor           =   ""
      NoCaracteres    =   "0123456789"
      SignodeMiles    =   -1  'True
      MarcarTextoAlEnfoque=   -1  'True
      NoRangoCadena   =   -1  'True
   End
End
Attribute VB_Name = "frmRepTransferencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
  Me.Width = 6000
  Me.Height = 1545
End Sub

Private Sub cmdAceptar_Click()
  Call Imprimir
End Sub

Private Sub cmdCancelar_Click()
  Unload Me
End Sub

Private Sub Imprimir()
 Dim rs As New ADODB.Recordset
 Dim SQL As String
 
 SQL = "Select cabrec_numrecibo from te_cabecerarecibos where cabrec_numreciboegreso='" & TxtTransf.Text & "'"
 Set rs = New ADODB.Recordset
 Set rs = VGcnx.Execute(SQL)
 If Not rs.BOF And Not rs.EOF Then
    Do Until rs.EOF
       Call ImprimirRecibo(rs(0))
       rs.MoveNext
    Loop
 Else
    MsgBox "No existe el Nro de Transferencia o ha sido Anulado", vbInformation, Caption
 End If
 rs.Close
 Set rs = Nothing
  
End Sub

Private Sub TxtTransf_LostFocus()
  TxtTransf.Text = Right(Format(TxtTransf.Text, "000000"), 6)
End Sub
