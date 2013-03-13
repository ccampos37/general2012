VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Frmseleccionfechatrabajo 
   Caption         =   "Form1"
   ClientHeight    =   1335
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3990
   LinkTopic       =   "Form1"
   ScaleHeight     =   1335
   ScaleWidth      =   3990
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdaceptar 
      Caption         =   "&Aceptar"
      Height          =   330
      Left            =   585
      TabIndex        =   4
      Top             =   825
      Width           =   1350
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   180
      TabIndex        =   1
      Top             =   135
      Width           =   3585
      Begin MSComCtl2.DTPicker DTPperiodo 
         Height          =   285
         Left            =   1830
         TabIndex        =   2
         Top             =   195
         Width           =   1650
         _ExtentX        =   2910
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "dd- MMM - yyyy"
         Format          =   16777219
         CurrentDate     =   37495
      End
      Begin VB.Label Label1 
         Caption         =   "Periodo de Trabajo :"
         Height          =   285
         Left            =   210
         TabIndex        =   3
         Top             =   225
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdcancelar 
      Caption         =   "&Cancelar"
      Height          =   330
      Left            =   2040
      TabIndex        =   0
      Top             =   825
      Width           =   1350
   End
End
Attribute VB_Name = "Frmseleccionfechatrabajo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdaceptar_Click()
Dim tccambio As Double
  Set VGvardllgen = New dllgeneral.dll_general
    VGParamSistem.AnoProceso = Format(Year(DTPperiodo), "0000")
    VGParamSistem.MesProceso = Format(Month(DTPperiodo), "00")
    VGParamSistem.fechatrabajo = DTPperiodo
    MDIPrincipal.Panel.Panels(1).Text = "Mes Proceso : " & VGvardllgen.DESMES(Month(DTPperiodo))
    MDIPrincipal.Panel.Panels(2).Text = "Año Proceso :" & Year(DTPperiodo)
    tccambio = XRecuperaTipoCambio(Format(DTPperiodo, "dd/mm/yyyy"), Venta, VGCnxCT)
    If tccambio = 0 Then
        MsgBox "No existe tipo de cambio para esta fecha", vbInformation
    End If
    MDIPrincipal.Panel.Panels(4).Text = "Tipo Cambio  (" & Format(tccambio, "#.000") & ")"
    MDIPrincipal.Panel.Panels(3).Text = "Fecha de Trabajo (" & VGParamSistem.fechatrabajo & ")"
    Unload Me
End Sub

Private Sub CmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
      DTPperiodo.Value = VGParamSistem.fechatrabajo
End Sub

