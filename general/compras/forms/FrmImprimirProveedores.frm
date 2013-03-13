VERSION 5.00
Begin VB.Form FrmImprimirProveedores 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte de proveedores"
   ClientHeight    =   1575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5910
   Icon            =   "FrmImprimirProveedores.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   5910
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdImp 
      Caption         =   "&Imprimir"
      Height          =   675
      Left            =   3720
      Picture         =   "FrmImprimirProveedores.frx":1E72
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   360
      Width           =   775
   End
   Begin VB.CommandButton CmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   675
      Left            =   4560
      Picture         =   "FrmImprimirProveedores.frx":22B4
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   360
      Width           =   775
   End
   Begin VB.Frame Frame3 
      Caption         =   "Listar Por "
      ForeColor       =   &H00FF8080&
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3255
      Begin VB.OptionButton OptExonerados 
         Caption         =   "Exonerados de Retencion"
         Height          =   735
         Left            =   1560
         TabIndex        =   2
         Top             =   120
         Width           =   1215
      End
      Begin VB.OptionButton OptTodos 
         Caption         =   "Todos"
         Height          =   495
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
   End
End
Attribute VB_Name = "FrmImprimirProveedores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ChkFech_Click()
If ChkFech.Value = 1 Then
    DTPFechaIni.Enabled = True
    DTPFechaFin.Enabled = True
  Else
    DTPFechaIni.Enabled = False
    DTPFechaFin.Enabled = False
End If
End Sub

Private Sub chkflagmodo_Click()
    If chkflagmodo.Value = 1 Then
        ChkCtaCte.Enabled = True: ChkActCaja.Enabled = True: ChkRegComp.Enabled = True
      Else
        ChkCtaCte.Enabled = False: ChkActCaja.Enabled = False: ChkRegComp.Enabled = False
    End If
End Sub


Private Sub cmdImp_Click()
Dim arrform(0) As Variant, arrparm(3) As Variant
On Error GoTo Imprime
arrparm(0) = VGCNx.DefaultDatabase
arrparm(1) = "cp_proveedor"
arrparm(2) = "isnull(proveedorcontribuyente,0)=1"
Call ImpresionRptProc("co_PrincipalContribuyente.rpt", arrform, arrparm, , "Reporte de Proveedores exonerados")
Screen.MousePointer = 1
Exit Sub
Imprime:
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Left = 2500
    Me.Top = 1400
    OptExonerados.Value = True
End Sub
