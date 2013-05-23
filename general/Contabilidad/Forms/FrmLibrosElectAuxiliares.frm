VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   8235
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10125
   LinkTopic       =   "Form2"
   ScaleHeight     =   8235
   ScaleWidth      =   10125
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   8055
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   9735
      Begin VB.Frame FrameCuentas 
         Height          =   6735
         Left            =   360
         TabIndex        =   1
         Top             =   960
         Width           =   9015
         Begin VB.Frame Frame2 
            Caption         =   "Opciones"
            Height          =   1215
            Left            =   2640
            TabIndex        =   2
            Top             =   5040
            Width           =   4215
            Begin VB.CommandButton Command1 
               Caption         =   "Impprimir"
               Height          =   615
               Left            =   360
               TabIndex        =   4
               Top             =   240
               Width           =   1575
            End
            Begin VB.CommandButton Command2 
               Caption         =   "Salir"
               Height          =   615
               Left            =   2400
               TabIndex        =   3
               Top             =   240
               Width           =   1575
            End
         End
         Begin MSComctlLib.ListView ListView2 
            Height          =   135
            Left            =   1320
            TabIndex        =   5
            Top             =   1560
            Width           =   30
            _ExtentX        =   53
            _ExtentY        =   238
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   4575
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   8415
            _ExtentX        =   14843
            _ExtentY        =   8070
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
