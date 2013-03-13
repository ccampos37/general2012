VERSION 5.00
Begin VB.Form FrmInventariosyBalances 
   Caption         =   "Inventarios y Balances"
   ClientHeight    =   3330
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7065
   LinkTopic       =   "Form1"
   ScaleHeight     =   3330
   ScaleWidth      =   7065
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   2655
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   6255
      Begin VB.OptionButton Option1 
         Caption         =   "Analitico"
         Height          =   495
         Left            =   720
         TabIndex        =   5
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Salir"
         Height          =   735
         Left            =   3360
         TabIndex        =   4
         Top             =   1320
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Imprimir"
         Height          =   735
         Left            =   1080
         TabIndex        =   3
         Top             =   1320
         Width           =   1575
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Resumido"
         Height          =   495
         Left            =   4560
         TabIndex        =   2
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Detalaldo"
         Height          =   495
         Left            =   2880
         TabIndex        =   1
         Top             =   360
         Width           =   1215
      End
   End
End
Attribute VB_Name = "FrmInventariosyBalances"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim aparm(5) As Variant
Dim aform(2) As Variant
aparm(0) = VGCNx.DefaultDatabase
aparm(1) = VGParametros.empresacodigo
aparm(2) = VGParamSistem.Anoproceso
aparm(3) = VGParamSistem.Mesproceso
aparm(4) = "1"

aform(0) = "@mes='" & VGvardllgen.DesMes(Trim$(VGParamSistem.Mesproceso)) & "'"
   
If Option1.Value = True Then
   aform(1) = "@tituloreporte='LIBRO INVENTARIOS Y BALANCE - ANALITICO '"
   Call ImpresionRptProc("ct_LibroInventariosXAnalitico.rpt", aform, aparm, , " Libro Inventarios y Balance ")
 ElseIf Option2.Value = True Then
    aparm(4) = "2"
    aform(1) = "@tituloreporte='LIBRO INVENTARIOS Y BALANCE - DETALLADO '"
    Call ImpresionRptProc("ct_LibroInventarios.rpt", aform, aparm, , " Libro Inventarios y Balance ")
  Else
   aform(1) = "@tituloreporte='LIBRO INVENTARIOS Y BALANCE - RESUMIDO '"
   Call ImpresionRptProc("ct_LibroInventarios.rpt", aform, aparm, , " Libro Inventarios y Balance ")
  End If
 
End Sub

Private Sub Command2_Click()
Unload Me
End Sub
