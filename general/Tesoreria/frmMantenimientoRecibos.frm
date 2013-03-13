VERSION 5.00
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "CONTROLAYUDA.OCX"
Begin VB.Form frmMantenimientoRecibos 
   Caption         =   "Mantenimiento Recibos"
   ClientHeight    =   4410
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5940
   LinkTopic       =   "Form1"
   ScaleHeight     =   4410
   ScaleWidth      =   5940
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   390
      Left            =   3075
      TabIndex        =   3
      Top             =   3795
      Width           =   1365
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   390
      Left            =   1605
      TabIndex        =   2
      Top             =   3795
      Width           =   1365
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos a Modificar"
      Height          =   2610
      Left            =   75
      TabIndex        =   1
      Top             =   945
      Width           =   5790
   End
   Begin ctrlayuda_f.Ctr_Ayuda Ctr_Recibos 
      Height          =   345
      Left            =   2250
      TabIndex        =   0
      Top             =   210
      Width           =   3300
      _ExtentX        =   5821
      _ExtentY        =   609
      Enabled         =   0   'False
      XcodMaxLongitud =   0
   End
   Begin VB.Label Label1 
      Caption         =   "Nº Recibo (I/E)"
      Height          =   285
      Left            =   315
      TabIndex        =   4
      Top             =   255
      Width           =   1305
   End
End
Attribute VB_Name = "frmMantenimientoRecibos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

