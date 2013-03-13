VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmrelacontra 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Relacion de Contratos"
   ClientHeight    =   555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4365
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   555
   ScaleWidth      =   4365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   315
      Left            =   3105
      TabIndex        =   2
      Top             =   90
      Width           =   1185
   End
   Begin MSComCtl2.DTPicker DTPperiodo 
      Height          =   315
      Left            =   1065
      TabIndex        =   1
      Top             =   90
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "MMM - yyyy "
      Format          =   24707075
      CurrentDate     =   37666
   End
   Begin VB.Label Label1 
      Caption         =   "Periodo : "
      Height          =   300
      Left            =   60
      TabIndex        =   0
      Top             =   105
      Width           =   840
   End
End
Attribute VB_Name = "frmrelacontra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Call IMPRIMIR
End Sub
Private Sub IMPRIMIR()
    Dim arrform(2) As Variant, arrparm(8) As Variant
    '@BASE, @ANO, @MES
    arrparm(0) = REGSISTEMA.BASESQL
    arrparm(1) = Format(Year(DTPperiodo.Value), "0000")
    arrparm(2) = Format(Month(DTPperiodo.Value), "0")
    Call ImpresionRptProc("pl_relacontra.rpt", arrform, arrparm, , "Relacion de Contratos")
End Sub
