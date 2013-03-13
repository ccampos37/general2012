VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmValidarComprobantes 
   Caption         =   "Errores en Comprobantes"
   ClientHeight    =   4680
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6150
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4680
   ScaleWidth      =   6150
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancelar"
      Height          =   360
      Left            =   3240
      TabIndex        =   9
      Top             =   4035
      Width           =   1410
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   360
      Left            =   1500
      TabIndex        =   8
      Top             =   4035
      Width           =   1350
   End
   Begin VB.Frame Frame2 
      Height          =   795
      Left            =   165
      TabIndex        =   3
      Top             =   675
      Width           =   5820
      Begin VB.OptionButton Opt 
         Caption         =   "Asientos Descuadrados"
         Height          =   375
         Left            =   165
         TabIndex        =   4
         Top             =   240
         Width           =   2130
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Validación de Cuentas en Comprobantes"
      Height          =   1950
      Left            =   150
      TabIndex        =   2
      Top             =   1725
      Width           =   5835
      Begin VB.OptionButton Option2 
         Caption         =   "Cruce de la Clase 6 con la 9"
         Height          =   360
         Left            =   180
         TabIndex        =   7
         Top             =   1215
         Width           =   2835
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Cruce de la Clase 6 con la 9"
         Height          =   360
         Left            =   180
         TabIndex        =   6
         Top             =   825
         Width           =   2835
      End
      Begin VB.OptionButton OptValidar 
         Caption         =   "Cruce de la Clase 6 con la 9"
         Height          =   360
         Left            =   180
         TabIndex        =   5
         Top             =   420
         Width           =   2835
      End
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   1815
      TabIndex        =   1
      Top             =   210
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   556
      _Version        =   393216
      CheckBox        =   -1  'True
      Format          =   24641537
      CurrentDate     =   37704
   End
   Begin VB.Label Label1 
      Caption         =   "Seleccionar Mes"
      Height          =   315
      Left            =   270
      TabIndex        =   0
      Top             =   255
      Width           =   1350
   End
End
Attribute VB_Name = "frmValidarComprobantes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

