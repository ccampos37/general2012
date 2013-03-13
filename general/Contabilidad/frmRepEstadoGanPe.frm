VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRepEstadosFinancieros 
   Caption         =   "Estados Financieros"
   ClientHeight    =   4290
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5895
   LinkTopic       =   "Form1"
   ScaleHeight     =   4290
   ScaleWidth      =   5895
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd 
      Caption         =   "Command1"
      Height          =   345
      Left            =   1230
      TabIndex        =   6
      Top             =   3315
      Width           =   1515
   End
   Begin VB.Frame Frame2 
      Caption         =   "Seleccionar el Mes"
      Height          =   810
      Left            =   15
      TabIndex        =   4
      Top             =   2115
      Width           =   5835
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   330
         Left            =   120
         TabIndex        =   5
         Top             =   270
         Width           =   3900
         _ExtentX        =   6879
         _ExtentY        =   582
         _Version        =   393216
         Format          =   24510465
         CurrentDate     =   37559
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Seleccionar Tipo"
      Height          =   1875
      Left            =   -15
      TabIndex        =   0
      Top             =   150
      Width           =   5865
      Begin VB.OptionButton optTipo 
         Caption         =   "Balance General"
         Height          =   480
         Index           =   2
         Left            =   120
         TabIndex        =   3
         Top             =   1230
         Width           =   4470
      End
      Begin VB.OptionButton optTipo 
         Caption         =   "Estado de Ganancias y Pérdidas por Naturaleza"
         Height          =   480
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   780
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

