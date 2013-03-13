VERSION 5.00
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "CONTROLAYUDA.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form RptCtactexVendedor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cuenta Corriente por Vendedor"
   ClientHeight    =   4275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5760
   Icon            =   "RptCtactexVendedor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   5760
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2955
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   5385
      Begin VB.ComboBox Combo4 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   2340
         Width           =   1245
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1920
         Width           =   1245
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1530
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
         Left            =   1680
         TabIndex        =   4
         Top             =   1110
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuda1 
         Height          =   285
         Left            =   1680
         TabIndex        =   5
         Top             =   330
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   503
         Enabled         =   0   'False
         XcodMaxLongitud =   0
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuda2 
         Height          =   285
         Left            =   1680
         TabIndex        =   6
         Top             =   720
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   503
         Enabled         =   0   'False
         XcodMaxLongitud =   0
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Incluido Letra x Abonar"
         Height          =   375
         Index           =   6
         Left            =   360
         TabIndex        =   12
         Top             =   2370
         Width           =   1185
      End
      Begin VB.Label Label1 
         Caption         =   "Moneda"
         Height          =   255
         Index           =   5
         Left            =   390
         TabIndex        =   11
         Top             =   1980
         Width           =   1185
      End
      Begin VB.Label Label1 
         Caption         =   "Con Resumen"
         Height          =   255
         Index           =   4
         Left            =   390
         TabIndex        =   10
         Top             =   1590
         Width           =   1245
      End
      Begin VB.Label Label1 
         Caption         =   "Hasta la Fecha"
         Height          =   255
         Index           =   2
         Left            =   390
         TabIndex        =   9
         Top             =   1170
         Width           =   1245
      End
      Begin VB.Label Label1 
         Caption         =   "Cliente"
         Height          =   255
         Index           =   1
         Left            =   390
         TabIndex        =   8
         Top             =   780
         Width           =   1245
      End
      Begin VB.Label Label1 
         Caption         =   "Vendedor"
         Height          =   255
         Index           =   0
         Left            =   390
         TabIndex        =   7
         Top             =   360
         Width           =   1245
      End
   End
End
Attribute VB_Name = "RptCtactexVendedor"
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
   Call CargarTipo(Combo4, 3)
   Combo3.Clear
   Combo3.AddItem g_TipoSol & "-Soles"
   Combo3.AddItem g_TipoDolar & "-Dolares"
   Combo3.ListIndex = 0
   MBox1.Text = Format(Date, "DD/MM/YYYY")
End Sub




