VERSION 5.00
Begin VB.Form frmRepgastosacumulados 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5940
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   5940
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Tipo de reporte"
      ForeColor       =   &H000000FF&
      Height          =   1155
      Left            =   3120
      TabIndex        =   5
      Top             =   120
      Width           =   2655
      Begin VB.OptionButton Option4 
         Caption         =   "Resumido"
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Detallado"
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ordenado Por : "
      ForeColor       =   &H000000FF&
      Height          =   1155
      Left            =   360
      TabIndex        =   2
      Top             =   120
      Width           =   2655
      Begin VB.OptionButton Option1 
         Caption         =   "Centro Costo / Cuenta"
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   360
         Width           =   2175
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Cuenta / Ccentro Costo"
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   720
         Width           =   2055
      End
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   360
      Left            =   1665
      TabIndex        =   1
      Top             =   1620
      Width           =   1275
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   360
      Left            =   3015
      TabIndex        =   0
      Top             =   1620
      Width           =   1275
   End
End
Attribute VB_Name = "frmRepgastosacumulados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAceptar_Click()
  Dim arrform(1) As Variant, arrparm(4) As Variant
  Set VGvardllgen = New dllgeneral.dll_general
  arrparm(0) = VGParamSistem.BDEmpresa
  arrparm(1) = VGParamSistem.Anoproceso
  arrparm(2) = VGParamSistem.Mesproceso
  
  arrform(0) = "@mes='" & VGvardllgen.DESMES(VGParamSistem.Mesproceso) & "'"
  
  If Option3.Value = True Then
       Call ImpresionRptProc("ct_cuentaxcentrodecostodetallado.rpt", arrform, arrparm, , "Reporte detallado ")
   ElseIf Option1.Value = True Then
          Call ImpresionRptProc("ct_cuentaxcentrocostoresumido.rpt", arrform, arrparm, , "Reporte resumido ")
        Else
          Call ImpresionRptProc("ct_listacentrocostoresumido.rpt", arrform, arrparm, , "Reporte resumido ")
   End If
   End Sub

Private Sub cmdCancelar_Click()
 Unload Me
End Sub

Private Sub Form_Load()
  Option1.Value = True
  Option3.Value = True
End Sub

