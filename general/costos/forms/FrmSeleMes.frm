VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmSeleMes 
   Caption         =   "Form1"
   ClientHeight    =   1425
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   4065
   LinkTopic       =   "Form1"
   ScaleHeight     =   1425
   ScaleWidth      =   4065
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdaceptar 
      Caption         =   "&Aceptar"
      Height          =   330
      Left            =   480
      TabIndex        =   4
      Top             =   840
      Width           =   1350
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   360
      TabIndex        =   1
      Top             =   120
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
         Format          =   20905987
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
      Left            =   2280
      TabIndex        =   0
      Top             =   840
      Width           =   1350
   End
End
Attribute VB_Name = "FrmSeleMes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAceptar_Click()
Dim tccambio As Double
  Set VGvardllgen = New dllgeneral.dll_general
    VGParamSistem.Anoproceso = Format(Year(DTPperiodo), "0000")
    VGParamSistem.Mesproceso = Format(Month(DTPperiodo), "00")
    VGParamSistem.FechaTrabajo = DTPperiodo
    MDIPrincipal.StatusBar1.Panels(1).Text = "Mes Proceso : " & VGvardllgen.DesMes(Month(DTPperiodo))
    MDIPrincipal.StatusBar1.Panels(2).Text = "Año Proceso :" & Year(DTPperiodo)
    tccambio = XRecuperaTipoCambio(Format(DTPperiodo, "dd/mm/yyyy"), Venta, VGCNx)
    If tccambio = 0 Then
        MsgBox "No existe tipo de cambio para esta fecha", vbInformation
    End If
    MDIPrincipal.StatusBar1.Panels(4).Text = "Tipo Cambio  (" & Format(tccambio, "#.000") & ")"
    MDIPrincipal.StatusBar1.Panels(3).Text = "Fecha de Trabajo (" & VGParamSistem.FechaTrabajo & ")"
    
    Unload Me
    
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
      DTPperiodo.Value = VGParamSistem.FechaTrabajo
End Sub

