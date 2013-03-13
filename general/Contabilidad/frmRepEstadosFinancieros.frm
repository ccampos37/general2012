VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRepEstadosFinancieros 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Estados Financieros"
   ClientHeight    =   4620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5895
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   5895
   Begin VB.CommandButton cmdBotones 
      Caption         =   "&Cancelar"
      Height          =   375
      Index           =   1
      Left            =   3097
      TabIndex        =   7
      Top             =   4065
      Width           =   1335
   End
   Begin VB.CommandButton cmdBotones 
      Caption         =   "&Aceptar"
      Height          =   375
      Index           =   0
      Left            =   1432
      TabIndex        =   6
      Top             =   4065
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Caption         =   "Seleccionar Período"
      Height          =   810
      Left            =   0
      TabIndex        =   4
      Top             =   2805
      Width           =   5865
      Begin MSComCtl2.DTPicker DTPicker_Fecha 
         Height          =   330
         Left            =   120
         TabIndex        =   5
         Top             =   270
         Width           =   3900
         _ExtentX        =   6879
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "MM - MMMM "
         Format          =   51445763
         UpDown          =   -1  'True
         CurrentDate     =   37559
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Seleccionar Estado Financiero"
      Height          =   2355
      Left            =   0
      TabIndex        =   0
      Top             =   180
      Width           =   5865
      Begin VB.Frame Frame3 
         Caption         =   "Tipo de Balance"
         Height          =   1095
         Left            =   3120
         TabIndex        =   9
         Top             =   1200
         Width           =   2295
         Begin VB.CheckBox Check1 
            Caption         =   "Diferenciado por Afiliadas"
            Height          =   495
            Left            =   240
            TabIndex        =   10
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.OptionButton optTipo 
         Caption         =   "Estados de Flujos de Efectivo"
         Height          =   480
         Index           =   3
         Left            =   120
         TabIndex        =   8
         Top             =   1680
         Width           =   4470
      End
      Begin VB.OptionButton optTipo 
         Caption         =   "Balance General"
         Height          =   480
         Index           =   2
         Left            =   120
         TabIndex        =   3
         Top             =   1230
         Width           =   2550
      End
      Begin VB.OptionButton optTipo 
         Caption         =   "Estado de Ganancias y Pérdidas por Naturaleza"
         Height          =   480
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   735
         Width           =   4470
      End
      Begin VB.OptionButton optTipo 
         Caption         =   "Estado de Ganancias y Pérdidas por Función"
         Height          =   480
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   255
         Width           =   4470
      End
   End
End
Attribute VB_Name = "frmRepEstadosFinancieros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    DTPicker_Fecha.Value = DateAdd("d", -1, DateAdd("m", 1, Format("01/" & VGParamSistem.Mesproceso & "/" & CInt(VGParamSistem.Anoproceso), "dd/mm/yyyy")))
    DTPicker_Fecha.MinDate = Format("01/01/" & VGParamSistem.Anoproceso, "dd/mm/yyyy")
    DTPicker_Fecha.MaxDate = Format("31/12/" & VGParamSistem.Anoproceso, "dd/mm/yyyy")
    Check1.Value = 0
    Frame3.Visible = False
End Sub

Sub ImpresionEFE()
'FIXIT: Declare 'arrform' and 'arrparm' con un tipo de datos de enlace en tiempo de compilación     FixIT90210ae-R1672-R1B8ZE
Dim arrform(0) As Variant, arrparm() As Variant
    ReDim arrparm(5)
    arrparm(0) = VGParamSistem.BDEmpresa
    arrparm(1) = VGParametros.empresacodigo
    arrparm(2) = VGParamSistem.Anoproceso
    arrparm(3) = Format(Month(DTPicker_Fecha.Value), "0#")
    arrparm(4) = VGComputer
    Call ImpresionRptProc("ct_EstadoFlujoEfectivo.rpt", arrform, arrparm)
End Sub

Sub ImpresionBalanceGeneral()
'FIXIT: Declare 'arrform' con un tipo de datos de enlace en tiempo de compilación          FixIT90210ae-R1672-R1B8ZE
Dim arrform(0) As Variant, arrparm(7) As Variant
    arrparm(0) = VGParamSistem.BDEmpresa
    arrparm(1) = VGParametros.empresacodigo
    arrparm(2) = VGParamSistem.Anoproceso
    arrparm(3) = Format(Month(DTPicker_Fecha.Value), "0#")
    arrparm(4) = "2"
    arrparm(5) = VGComputer
    arrparm(6) = Check1.Value
    Call ImpresionRptProc("ct_BalanceGeneral.rpt", arrform, arrparm)
    Call ImpresionRptProc("ct_balancegeneral_anexo.rpt", arrform, arrparm)
End Sub

Sub ImpresionEGPFuncion()
Dim arrform(1) As Variant, arrparm(6) As Variant
    arrparm(0) = VGParamSistem.BDEmpresa
    arrparm(1) = VGParametros.empresacodigo
    arrparm(2) = VGParamSistem.Anoproceso
    arrparm(3) = Format(Month(Format(DTPicker_Fecha.Value, "dd/mm/yyyy")), "0#")
    arrparm(4) = VGComputer
   arrparm(5) = "0"
    arrform(0) = "@TituloReporte='" & "Estado de Ganáncias y Pérdidas por FUNCION" & "'"
    Call ImpresionRptProc("ct_EstadoGanPedFun.rpt", arrform, arrparm)
    End Sub

Sub ImpresionEGPNaturaleza()
Dim arrform(1) As Variant, arrparm(5) As Variant
    arrparm(0) = VGCNx.DefaultDatabase
    arrparm(1) = VGParametros.empresacodigo
    arrparm(2) = VGParamSistem.Anoproceso
    arrparm(3) = Format(Month(Format(DTPicker_Fecha.Value, "dd/mm/yyyy")), "0#")
    arrparm(4) = VGComputer
    arrform(0) = "@TituloReporte='" & "Estado de Ganáncias y Pérdidas por NATURALEZA" & "'"
    Call ImpresionRptProc("ct_EstadoGanPedNat.rpt", arrform, arrparm)
End Sub

Private Sub cmdBotones_Click(Index As Integer)

  If Index = 0 Then
    If optTipo(0).Value = True Then
           Call ImpresionEGPFuncion
    ElseIf optTipo(1).Value = True Then
           Call ImpresionEGPNaturaleza
    ElseIf optTipo(2) = True Then
           Call ImpresionBalanceGeneral
    Else
           Call ImpresionEFE
    End If
  Else
    Unload Me
  End If
End Sub

Private Sub optTipo_Click(Index As Integer)
Frame3.Visible = False
If Index = 2 Then Frame3.Visible = True
End Sub
