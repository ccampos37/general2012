VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form RptResumenCobranzaDetallada 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Resumen de Cobranza Detallada"
   ClientHeight    =   3210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3675
   Icon            =   "RptResumenCobranzaDetallada.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   3675
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2115
      Left            =   330
      TabIndex        =   0
      Top             =   150
      Width           =   2805
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1350
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1560
         Width           =   1245
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1350
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1110
         Width           =   1245
      End
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
         TabIndex        =   2
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
         TabIndex        =   3
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
         Caption         =   "Moneda"
         Height          =   255
         Index           =   0
         Left            =   390
         TabIndex        =   10
         Top             =   1590
         Width           =   825
      End
      Begin VB.Label Label1 
         Caption         =   "Resumen"
         Height          =   255
         Index           =   5
         Left            =   360
         TabIndex        =   6
         Top             =   1170
         Width           =   945
      End
      Begin VB.Label Label1 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   2
         Left            =   390
         TabIndex        =   5
         Top             =   750
         Width           =   585
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
   End
   Begin VB.PictureBox SSFrame1 
      Height          =   645
      Left            =   690
      ScaleHeight     =   585
      ScaleWidth      =   2145
      TabIndex        =   7
      Top             =   2370
      Width           =   2205
      Begin VB.PictureBox cAcepta 
         Height          =   375
         Left            =   120
         ScaleHeight     =   315
         ScaleWidth      =   885
         TabIndex        =   8
         Top             =   180
         Width           =   945
      End
      Begin VB.PictureBox cCancela 
         Height          =   375
         Left            =   1140
         ScaleHeight     =   315
         ScaleWidth      =   885
         TabIndex        =   9
         Top             =   180
         Width           =   945
      End
   End
End
Attribute VB_Name = "RptResumenCobranzaDetallada"
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
   Call CargarTipo(Combo1, 3)
   Combo2.Clear
   Combo2.AddItem g_TipoSol & "-Soles"
   Combo2.AddItem g_TipoDolar & "-Dolares"
   Combo2.ListIndex = 0
   MBox1.Text = Format(Date, "DD/MM/YYYY")
   MBox2.Text = Format(Date, "DD/MM/YYYY")
End Sub







