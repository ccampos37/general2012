VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form RptResumenCobranzaDiaria 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Resumen de Cobranza Diaria"
   ClientHeight    =   2595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3525
   Icon            =   "RptResumenCobranzaDiaria.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2595
   ScaleWidth      =   3525
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1275
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2895
      Begin MSMask.MaskEdBox MBox1 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "d/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   3
         EndProperty
         Height          =   315
         Left            =   1350
         TabIndex        =   1
         Top             =   270
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MBox2 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "d/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   3
         EndProperty
         Height          =   315
         Left            =   1350
         TabIndex        =   2
         Top             =   690
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         Caption         =   "Desde"
         Height          =   255
         Index           =   1
         Left            =   390
         TabIndex        =   4
         Top             =   330
         Width           =   1245
      End
      Begin VB.Label Label1 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   2
         Left            =   390
         TabIndex        =   3
         Top             =   780
         Width           =   585
      End
   End
   Begin VB.PictureBox SSFrame1 
      Height          =   645
      Left            =   720
      ScaleHeight     =   585
      ScaleWidth      =   2145
      TabIndex        =   5
      Top             =   1650
      Width           =   2205
      Begin VB.PictureBox cAcepta 
         Height          =   375
         Left            =   120
         ScaleHeight     =   315
         ScaleWidth      =   885
         TabIndex        =   6
         Top             =   180
         Width           =   945
      End
      Begin VB.PictureBox cCancela 
         Height          =   375
         Left            =   1140
         ScaleHeight     =   315
         ScaleWidth      =   885
         TabIndex        =   7
         Top             =   180
         Width           =   945
      End
   End
End
Attribute VB_Name = "RptResumenCobranzaDiaria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adll As New dllgeneral.dll_general


Private Sub cCancela_Click()
  Unload Me
End Sub

Private Sub Form_Load()
   MostrarForm Me, "C2"

   MBox1.Text = Format(Date, "DD/MM/YYYY")
   MBox2.Text = Format(Date, "DD/MM/YYYY")
End Sub






