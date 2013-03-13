VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmRepComprobantesRetencion 
   Caption         =   "Reporte de Comprobantes de Retencion Emitidos"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6075
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   6075
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBotones 
      Caption         =   "&Aceptar"
      Height          =   360
      Index           =   0
      Left            =   1545
      TabIndex        =   5
      Top             =   2100
      Width           =   1440
   End
   Begin VB.CommandButton cmdBotones 
      Caption         =   "&Cancelar"
      Height          =   360
      Index           =   1
      Left            =   2910
      TabIndex        =   4
      Top             =   2100
      Width           =   1440
   End
   Begin MSComCtl2.DTPicker DTPickerFecFinal 
      Height          =   300
      Left            =   3735
      TabIndex        =   0
      Top             =   525
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   529
      _Version        =   393216
      Format          =   25493505
      CurrentDate     =   37474
   End
   Begin MSComCtl2.DTPicker DTPickerFecInicio 
      Height          =   300
      Left            =   1125
      TabIndex        =   1
      Top             =   525
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   529
      _Version        =   393216
      Format          =   25493505
      CurrentDate     =   37474
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha Inicial"
      Height          =   300
      Left            =   165
      TabIndex        =   3
      Top             =   570
      Width           =   930
   End
   Begin VB.Label Label2 
      Caption         =   "Fecha Final"
      Height          =   285
      Left            =   2835
      TabIndex        =   2
      Top             =   570
      Width           =   840
   End
End
Attribute VB_Name = "frmRepComprobantesRetencion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
  Dim cFecha As Date
 
  DTPickerFecInicio.Value = Format("01/" & Format(Month(Date), "00") & "/" & Year(Date), "dd/mm/yyyy")
  cFecha = Format("01/" & Format(Month(Date) + 1, "00") & "/" & Year(Date), "dd/mm/yyyy")
  DTPickerFecFinal.Value = Format(cFecha - 1, "dd/mm/yyyy")
End Sub

Private Sub cmdBotones_Click(Index As Integer)
  Select Case Index
    Case 0
       Call ImpresionComprobantesEmitidos
    Case 1
       Unload Me
  End Select
   
End Sub

Sub ImpresionComprobantesEmitidos()
Dim arrform() As Variant, arrparm() As Variant
Dim dato As String
ReDim arrparm(5)
ReDim arrform(1)

arrparm(0) = VGParamSistem.BDEmpresa
arrparm(1) = "%%"
arrparm(2) = VGParametros.empresacodigoretencion
arrparm(3) = Format(DTPickerFecInicio.Value, "dd/mm/yyyy")
arrparm(4) = Format(DTPickerFecFinal.Value, "dd/mm/yyyy")

dato = "Del : " & DTPickerFecInicio & " Al : " & DTPickerFecFinal
arrform(0) = "dato='" & dato & "'"

Call ImpresionRptProc("te_ListadoComprobantesRetencion.rpt ", arrform, arrparm)

End Sub

