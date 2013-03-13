VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmselanomes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Seleccion Periodo de Trabajo"
   ClientHeight    =   1065
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3720
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1065
   ScaleWidth      =   3720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdcancelar 
      Caption         =   "&Cancelar"
      Height          =   330
      Left            =   1905
      TabIndex        =   3
      Top             =   660
      Width           =   1350
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   45
      TabIndex        =   1
      Top             =   -30
      Width           =   3585
      Begin MSComCtl2.DTPicker DTPperiodo 
         Height          =   285
         Left            =   1830
         TabIndex        =   4
         Top             =   195
         Width           =   1650
         _ExtentX        =   2910
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "dd- MMM - yyyy"
         Format          =   58130435
         CurrentDate     =   37495
      End
      Begin VB.Label Label1 
         Caption         =   "Periodo de Trabajo :"
         Height          =   285
         Left            =   210
         TabIndex        =   2
         Top             =   225
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdaceptar 
      Caption         =   "&Aceptar"
      Height          =   330
      Left            =   450
      TabIndex        =   0
      Top             =   660
      Width           =   1350
   End
End
Attribute VB_Name = "frmselanomes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdAceptar_Click()
Dim tccambio As Double
  Set VGvardllgen = New dllgeneral.dll_general
    VGParamSistem.Anoproceso = Format(Year(DTPperiodo), "0000")
    VGParamSistem.Mesproceso = Format(Month(DTPperiodo), "00")
    VGParamSistem.FechaTrabajo = DTPperiodo
    MDIPrincipal.StatusBar1.Panels(1).Text = "Mes Proceso : " & VGvardllgen.DESMES(Month(DTPperiodo))
    MDIPrincipal.StatusBar1.Panels(2).Text = "Año Proceso :" & Year(DTPperiodo)
    
    tccambio = XRecuperaTipoCambio(Format(DTPperiodo, "dd/mm/yyyy"), Venta, VGcnxCT)
    If tccambio = 0 Then
        MsgBox "No existe tipo de cambio para esta fecha", vbInformation
    End If
    MDIPrincipal.StatusBar1.Panels(4).Text = "Tipo Cambio  (" & Format(tccambio, "#.000") & ")"
    MDIPrincipal.StatusBar1.Panels(3).Text = "Fecha de Trabajo (" & VGParamSistem.FechaTrabajo & ")"
    
    Unload Me
    
End Sub

Private Sub CmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
      DTPperiodo.Value = VGParamSistem.FechaTrabajo
End Sub
