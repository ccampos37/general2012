VERSION 5.00
Object = "{4D137D9C-00A6-4458-9B46-7E95DB76D55B}#9.0#0"; "TextFer.ocx"
Begin VB.Form FrmRepOrdCompra 
   Caption         =   "Impresion de Orden de Compra"
   ClientHeight    =   915
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5085
   LinkTopic       =   "Form1"
   ScaleHeight     =   915
   ScaleWidth      =   5085
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdcancelar 
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   3345
      TabIndex        =   3
      Top             =   495
      Width           =   1350
   End
   Begin VB.CommandButton cmdaceptar 
      Caption         =   "&Aceptar"
      Height          =   330
      Left            =   3345
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin TextFer.TxFer TxtNroOrdenCompra 
      Height          =   300
      Left            =   90
      TabIndex        =   1
      Top             =   495
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   529
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
      NoCaracteres    =   "0123456789"
      NoRangoCadena   =   -1  'True
   End
   Begin VB.Label Label1 
      Caption         =   "Nro de Orden de Compra :"
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   2640
   End
End
Attribute VB_Name = "FrmRepOrdCompra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdaceptar_Click()
    Call imprimir
End Sub

Private Sub CmdCancelar_Click()
    Unload Me
End Sub
Public Sub imprimir()
Dim arrform(0) As Variant, arrparm(13) As Variant
Dim rsaux As ADODB.Recordset
On Error GoTo Imprime
    Set rsaux = New ADODB.Recordset
    
    'Val(FrmOrdenCompra.LblParte.Caption) & ", " & Val(FrmOrdenCompra.Tipcom)
    rsaux.Open "select ctipo=isnull(CompraTipo,0) from dbo.OrdenCompra " & _
               "where OrdenNro=" & Trim(TxtNroOrdenCompra.Text), VGCNx, adOpenKeyset, adLockReadOnly
    
    If rsaux.RecordCount = 0 Then
        MsgBox "No existe la orden de compra solicitada ", vbExclamation
        Exit Sub
    End If
    Screen.MousePointer = 11
    Dim rsdate As New ADODB.Recordset
    Set VGvardllgen = New dllgeneral.dll_general
    arrparm(0) = Val(Trim(TxtNroOrdenCompra.Text))
    arrparm(1) = rsaux!ctipo
    Call ImpresionRptbase("rptcoestaordencompra.rpt", arrform, arrparm, , "Listado de orden de compra pendientes ")
    Screen.MousePointer = 1
    Exit Sub
Imprime:
Screen.MousePointer = 1
MsgBox Err.Description
End Sub
