VERSION 5.00
Begin VB.Form FrmLibroCajayBancos 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Imprimir Libro Caja y Bancos"
   ClientHeight    =   4245
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6720
   LinkTopic       =   "Form1"
   ScaleHeight     =   4245
   ScaleWidth      =   6720
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFC0&
      Height          =   1215
      Left            =   1200
      TabIndex        =   3
      Top             =   2280
      Width           =   3975
      Begin VB.CommandButton Command2 
         BackColor       =   &H0000C000&
         Caption         =   "Salir"
         Height          =   495
         Left            =   2160
         TabIndex        =   5
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Imprimir"
         Height          =   495
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Height          =   1695
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   5535
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Detalle de los movimeintos de cuenta corriente"
         Height          =   495
         Left            =   480
         TabIndex        =   2
         Top             =   840
         Width           =   3855
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Detalle de los movimientos de efectivo"
         Height          =   495
         Left            =   480
         TabIndex        =   1
         Top             =   360
         Width           =   3615
      End
   End
End
Attribute VB_Name = "FrmLibroCajayBancos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Dim aparam(9) As Variant
Dim aform(1) As Variant
Dim RSQL As New ADODB.Recordset
aform(0) = "empresa='" & VGParametros.NomEmpresa & "'"
aparam(0) = VGCNx.DefaultDatabase
aparam(1) = VGParametros.empresacodigo
aparam(2) = VGParamSistem.Anoproceso
aparam(3) = VGParamSistem.Mesproceso
aparam(4) = Format(VGParamSistem.Mesproceso - 1, "00")
aparam(7) = VGParametros.sistemactaajustedeb
aparam(8) = VGParametros.sistemactaajustehab
If Check1.Value = 1 Then
    aparam(5) = "FORMATO 01.01"
    Set RSQL = VGCNx.Execute("select formatocuentacomodin from ct_formatos where formatocodigo='" & aparam(5) & "'")
    aparam(6) = RSQL!formatocuentacomodin
    Call ImpresionRptProc("ct_LibroCajayBancos0101.rpt", aform, aparam, , Check1.Caption)
End If
If Check2.Value = 1 Then
    aparam(5) = "FORMATO 01.02"
    Set RSQL = VGCNx.Execute("select formatocuentacomodin from ct_formatos where formatocodigo='" & aparam(5) & "'")
    aparam(6) = RSQL!formatocuentacomodin
    Call ImpresionRptProc("ct_LibroCajayBancos0102.rpt", aform, aparam, , Check2.Caption)
End If

End Sub

Private Sub Command2_Click()
Unload Me
End Sub
