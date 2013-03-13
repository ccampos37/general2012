VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C5FF36B7-A67E-11D3-9D9D-E6F193F7F854}#9.0#0"; "ApliCTxt.ocx"
Begin VB.Form frAceptaCTS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Aceptar Planilla de C.T.S."
   ClientHeight    =   3285
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frAceptaCTS.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   315
      Left            =   2573
      TabIndex        =   6
      Top             =   2790
      Width           =   1215
   End
   Begin VB.CommandButton cmAceptar 
      Caption         =   "&Aceptar"
      Height          =   315
      Left            =   885
      TabIndex        =   9
      Top             =   2790
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Aceptar Planilla"
      Height          =   1875
      Left            =   120
      TabIndex        =   0
      Top             =   765
      Width           =   4455
      Begin AplisetControlText.Aplitext xDolar 
         Height          =   315
         Left            =   2265
         TabIndex        =   10
         Top             =   1230
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   556
         Text            =   "0"
         TipoDato        =   "N"
      End
      Begin MSComCtl2.DTPicker xFecha 
         Height          =   300
         Left            =   2295
         TabIndex        =   2
         Top             =   450
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   529
         _Version        =   393216
         Format          =   50003969
         CurrentDate     =   36826
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Precio del Dolar"
         Height          =   195
         Left            =   210
         TabIndex        =   5
         Top             =   1320
         Width           =   1125
      End
      Begin VB.Label xMonto 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00 "
         Height          =   300
         Left            =   2280
         TabIndex        =   4
         Top             =   855
         Width           =   1395
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Monto Total del Pago"
         Height          =   195
         Left            =   225
         TabIndex        =   3
         Top             =   908
         Width           =   1530
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha de Pago de Planilla"
         Height          =   195
         Left            =   210
         TabIndex        =   1
         Top             =   510
         Width           =   1860
      End
   End
   Begin VB.Label xPlanilla 
      BackColor       =   &H8000000C&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   120
      TabIndex        =   8
      Top             =   345
      Width           =   4455
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Planilla de CTS"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   1110
   End
End
Attribute VB_Name = "frAceptaCTS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMACEPTAR_CLICK()
    If MsgBox("ESTA SEGURO DE ACEPTAR LOS VALORES PARA LA PLANILLA DE CTS " & xPlanilla.Caption, vbYesNo + vbQuestion) = vbNo Then Exit Sub
    DBSYSTEM.Execute "UPDATE CTS SET CERRADO=1, DOLARES=" & Val(xMonto.Caption) * Val(xDolar.Text) & " WHERE CODIGO=" & VPTRASPRM
    DBSYSTEM.Execute "UPDATE PLANCTS SET FECHADEPOSITO=" & DateSQL(xFecha.Value) & " WHERE CODIGO=" & VPTRASPRM
    Unload Me
End Sub

Private Sub CMCANCELAR_CLICK()
    Unload Me
End Sub

Private Sub Form_Load()
    xPlanilla.Caption = DevuelveValor("SELECT NOMBRE FROM CTS WHERE CODIGO=" & VPTRASPRM, DBSYSTEM)
    xMonto.Caption = DevuelveValor("SELECT SOLES FROM CTS WHERE CODIGO=" & VPTRASPRM, DBSYSTEM)
    xDolar.Text = MDIPrincipal.BarraEstado.Panels("Dolar").Text
End Sub

