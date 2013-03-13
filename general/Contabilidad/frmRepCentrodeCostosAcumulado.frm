VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmRepCentrodeCostosAcumulado 
   Caption         =   "Acumulados de Cuentas de gastos"
   ClientHeight    =   2880
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5820
   LinkTopic       =   "Form1"
   ScaleHeight     =   2880
   ScaleWidth      =   5820
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   360
      Left            =   2745
      TabIndex        =   10
      Top             =   2250
      Width           =   1275
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   360
      Left            =   1395
      TabIndex        =   9
      Top             =   2250
      Width           =   1275
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ordenado Por : "
      ForeColor       =   &H000000FF&
      Height          =   1155
      Left            =   3015
      TabIndex        =   6
      Top             =   1035
      Width           =   2655
      Begin VB.OptionButton Option2 
         Caption         =   "Centro Costo"
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   720
         Width           =   2055
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Cuenta / Ccentro Costo"
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tipo de reporte"
      ForeColor       =   &H000000FF&
      Height          =   1155
      Left            =   105
      TabIndex        =   3
      Top             =   990
      Width           =   2655
      Begin VB.OptionButton Option3 
         Caption         =   "Mensual/Acumulado"
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   360
         Width           =   1815
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Mensualizado/Anual"
         Height          =   255
         Left            =   405
         TabIndex        =   4
         Top             =   720
         Width           =   1935
      End
   End
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   90
      TabIndex        =   0
      Top             =   270
      Width           =   5415
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   285
         Left            =   1410
         TabIndex        =   1
         Top             =   240
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "MM - MMMM"
         Format          =   25034755
         UpDown          =   -1  'True
         CurrentDate     =   37505
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Mes :"
         Height          =   195
         Left            =   480
         TabIndex        =   2
         Top             =   255
         Width           =   390
      End
   End
End
Attribute VB_Name = "frmRepCentrodeCostosAcumulado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAceptar_Click()
If Option3.Value = True Then
   Call imprimir3
ElseIf Option4.Value = True Then
   Call imprimir4
End If
End Sub
Private Sub imprimir3()
  Dim arrform(2) As Variant, arrparm(4) As Variant
  Set VGvardllgen = New dllgeneral.dll_general
  arrparm(0) = VGParamSistem.BDEmpresa
  arrparm(1) = VGParametros.empresacodigo
  arrparm(2) = VGParamSistem.Anoproceso
  arrparm(3) = Format(Month(DTPicker1.Value), "00")

  arrform(0) = "@mes='" & VGvardllgen.DesMes(VGParamSistem.Mesproceso) & "'"
  arrform(1) = "@Tituloreporte='Acumulados por Centro de Costos '"

     If Option1.Value = True Then
        Call ImpresionRptProc("ct_listacentrocostosresumen.rpt", arrform, arrparm, , "Reporte Centro de Costos x Cuentas Acumulado ")
       Else
        Call ImpresionRptProc("ct_cuentaxcentrocostodetallado.rpt", arrform, arrparm, , "Reporte acumulado ")
     End If
End Sub

Private Sub imprimir4()
  Dim arrform(2) As Variant, arrparm(5) As Variant
  Set VGvardllgen = New dllgeneral.dll_general
  arrparm(0) = VGParamSistem.BDEmpresa
  arrparm(1) = VGParametros.empresacodigo
  arrparm(2) = VGParamSistem.Anoproceso
  arrparm(3) = Format(Month(DTPicker1.Value), "00")

  arrform(0) = "@mes='" & VGvardllgen.DesMes(VGParamSistem.Mesproceso) & "'"
  arrform(1) = "@Tituloreporte='Acumulados por Centro de Costos '"

   If Option1.Value = True Then
          arrparm(4) = 0
          Call ImpresionRptProc("ct_listacentrocostoAcumMes.rpt", arrform, arrparm, , "Reporte resumido ")
       Else
        arrparm(4) = 1
          Call ImpresionRptProc("ct_listacentrocostoAcumxMeses.rpt", arrform, arrparm, , "Reporte resumido ")
     End If
 End Sub

Private Sub cmdCancelar_Click()
 Unload Me
End Sub

Private Sub Form_Load()
  Option1.Value = True
  Option3.Value = True
  
    DTPicker1.Value = DateSerial(CInt(VGParamSistem.Anoproceso), CInt(VGParamSistem.Mesproceso), 1)

End Sub


Private Sub Option3_Click()
If Option3.Value = True Then
   Option2.Enabled = False
End If
End Sub

Private Sub Option4_Click()
If Option4.Value = True Then
   Option2.Enabled = True
End If
End Sub
