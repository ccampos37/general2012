VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form RptPlanillaCobranza 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Planilla de Cobranza"
   ClientHeight    =   3210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5400
   Icon            =   "RptPlanillaCobranza.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   5400
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1875
      Left            =   270
      TabIndex        =   0
      Top             =   210
      Width           =   4905
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1200
         Width           =   1245
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   750
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
         TabIndex        =   3
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
         Left            =   3450
         TabIndex        =   8
         Top             =   270
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         Caption         =   "Con Detalle"
         Height          =   255
         Index           =   5
         Left            =   420
         TabIndex        =   7
         Top             =   1200
         Width           =   1185
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo Vendedor"
         Height          =   255
         Index           =   4
         Left            =   390
         TabIndex        =   6
         Top             =   780
         Width           =   1245
      End
      Begin VB.Label Label1 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   2
         Left            =   2790
         TabIndex        =   5
         Top             =   300
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
End
Attribute VB_Name = "RptPlanillaCobranza"
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
   Call CargarTipo(Combo2, 3)
   Combo1.Clear
   Combo1.AddItem "V-Vendedor"
   Combo1.AddItem "C-Caja"
   Combo1.AddItem "B-Dietario Bancos"
   Combo1.AddItem "T-Todos"
   Combo1.ListIndex = 0
   MBox1.Text = Format(Date, "DD/MM/YYYY")
   MBox2.Text = Format(Date, "DD/MM/YYYY")
End Sub





