VERSION 5.00
Begin VB.Form frCfgProc 
   Caption         =   "Cálculo de Planilla"
   ClientHeight    =   2415
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frCfgProc.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2415
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&Continuar"
      Default         =   -1  'True
      Height          =   360
      Left            =   698
      TabIndex        =   6
      Top             =   1815
      Width           =   1470
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   360
      Left            =   2513
      TabIndex        =   5
      Top             =   1815
      Width           =   1470
   End
   Begin VB.CheckBox xNuevoAdel 
      Caption         =   "Cargar Nuevos Adelantos"
      Height          =   210
      Left            =   180
      TabIndex        =   4
      Top             =   1350
      Width           =   4350
   End
   Begin VB.CheckBox Ing 
      Caption         =   "Asignar los ingresos pendientes"
      Height          =   225
      Left            =   180
      TabIndex        =   3
      Top             =   1080
      Width           =   2550
   End
   Begin VB.CheckBox Quinta 
      Caption         =   "Calcular impuesto de Quinta Categoria"
      Height          =   240
      Left            =   180
      TabIndex        =   2
      Top             =   840
      Width           =   3990
   End
   Begin VB.CheckBox Pres 
      Caption         =   "Descontar Prestamos y otros Descuentos pendientes"
      Height          =   195
      Left            =   180
      TabIndex        =   1
      Top             =   600
      Width           =   4230
   End
   Begin VB.CheckBox Adel 
      Caption         =   "Descontar Adelantos pendientes"
      Height          =   210
      Left            =   180
      TabIndex        =   0
      Top             =   345
      Width           =   3960
   End
End
Attribute VB_Name = "frCfgProc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub COMMAND1_CLICK()
    With RegProc
        .Adelantos = IIf(Adel.Value = 0, False, True)
        .Continuar = True
        .INGRESOS = IIf(Ing.Value = 0, False, True)
        .Prestamos = IIf(Pres.Value = 0, False, True)
        .Quinta = IIf(Quinta.Value = 0, False, True)
        .NuevosAdel = IIf(xNuevoAdel.Value = 0, False, True)
    End With
    Unload Me
End Sub

Private Sub COMMAND2_CLICK()
    RegProc.Continuar = False
    Unload Me
End Sub

Private Sub FORM_LOAD()
    RegProc.Continuar = False
End Sub
