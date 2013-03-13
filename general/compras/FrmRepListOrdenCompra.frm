VERSION 5.00
Begin VB.Form FrmRepListOrdenCompra 
   Caption         =   "Listado de Ordenes de Compra por Estado"
   ClientHeight    =   1530
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4830
   LinkTopic       =   "Form1"
   ScaleHeight     =   1530
   ScaleWidth      =   4830
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   2940
      TabIndex        =   5
      Top             =   735
      Width           =   1290
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   345
      Left            =   2940
      TabIndex        =   4
      Top             =   330
      Width           =   1290
   End
   Begin VB.Frame Frame1 
      Caption         =   "Estado"
      Height          =   1110
      Left            =   195
      TabIndex        =   0
      Top             =   120
      Width           =   2175
      Begin VB.OptionButton OptEstado 
         Caption         =   "Pendiente"
         Height          =   270
         Index           =   0
         Left            =   165
         TabIndex        =   3
         Top             =   240
         Width           =   1845
      End
      Begin VB.OptionButton OptEstado 
         Caption         =   "Parcialm. Atendido"
         Height          =   270
         Index           =   1
         Left            =   165
         TabIndex        =   2
         Top             =   525
         Width           =   1845
      End
      Begin VB.OptionButton OptEstado 
         Caption         =   "Atendido"
         Height          =   270
         Index           =   2
         Left            =   180
         TabIndex        =   1
         Top             =   810
         Width           =   1845
      End
   End
End
Attribute VB_Name = "FrmRepListOrdenCompra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vestado As Integer
Dim vdescr As String
Private Sub cmdaceptar_Click()
    Call imprimir
End Sub
Public Sub imprimir()
Dim arrform(1) As Variant, arrparm(3) As Variant
On Error GoTo Imprime
    Screen.MousePointer = 11
    Dim rsdate As New ADODB.Recordset
    Set VGvardllgen = New dllgeneral.dll_general
    arrform(0) = "esta='" & vdescr & "'"
    arrparm(0) = VGCNx.DefaultDatabase
    arrparm(1) = vestado
    Call ImpresionRptbase("rptCoEstadOrdenCompra.rpt", arrform, arrparm, , "Listado de orden de compra pendientes ")
    Screen.MousePointer = 1
    Exit Sub
Imprime:
    Screen.MousePointer = 1
    MsgBox Err.Description
End Sub


Private Sub OptEstado_Click(Index As Integer)
    vestado = Index + 1
    vdescr = UCase(OptEstado(Index).Caption)
End Sub
