VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form2 
   Caption         =   "Fecha"
   ClientHeight    =   2115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3915
   ClipControls    =   0   'False
   Icon            =   "FrmFecha.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   3915
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   3495
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   285
         Left            =   1560
         TabIndex        =   2
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         Format          =   24576001
         CurrentDate     =   36437
         MaxDate         =   401768
         MinDate         =   36161
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   285
         Left            =   1560
         TabIndex        =   3
         Top             =   840
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         Format          =   24576001
         CurrentDate     =   36437
         MaxDate         =   401768
         MinDate         =   36161
      End
      Begin VB.Label Label5 
         Caption         =   "Desde"
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Hasta"
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   840
         Width           =   615
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Sigue"
      Height          =   375
      Left            =   2640
      TabIndex        =   0
      Top             =   1680
      Width           =   1095
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
