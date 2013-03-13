VERSION 5.00
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "CONTROLAYUDA.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Rptclientexzona 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Clientes por Zonas"
   ClientHeight    =   2940
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4665
   Icon            =   "Rptclientexzona.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   4665
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Height          =   1695
      Left            =   210
      TabIndex        =   3
      Top             =   240
      Width           =   4245
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1260
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1110
         Width           =   2655
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuda1 
         Height          =   315
         Left            =   1260
         TabIndex        =   5
         Top             =   630
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   556
         Enabled         =   0   'False
         XcodMaxLongitud =   0
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
         Left            =   1830
         TabIndex        =   8
         Top             =   240
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Referencial"
         Height          =   255
         Index           =   1
         Left            =   270
         TabIndex        =   9
         Top             =   270
         Width           =   1365
      End
      Begin VB.Label Label2 
         Caption         =   "Ordenado"
         Height          =   165
         Index           =   0
         Left            =   270
         TabIndex        =   7
         Top             =   1170
         Width           =   1125
      End
      Begin VB.Label Label2 
         Caption         =   "Zona"
         Height          =   165
         Index           =   2
         Left            =   270
         TabIndex        =   6
         Top             =   720
         Width           =   1125
      End
   End
   Begin VB.PictureBox SSFrame1 
      Height          =   645
      Left            =   1260
      ScaleHeight     =   585
      ScaleWidth      =   2145
      TabIndex        =   0
      Top             =   2070
      Width           =   2205
      Begin VB.PictureBox cAcepta 
         Height          =   375
         Left            =   120
         ScaleHeight     =   315
         ScaleWidth      =   885
         TabIndex        =   1
         Top             =   180
         Width           =   945
      End
      Begin VB.PictureBox cCancela 
         Height          =   375
         Left            =   1140
         ScaleHeight     =   315
         ScaleWidth      =   885
         TabIndex        =   2
         Top             =   180
         Width           =   945
      End
   End
End
Attribute VB_Name = "Rptclientexzona"
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
   Combo1.Clear
   Combo1.AddItem "C-Codigo"
   Combo1.AddItem "D-Descripcion"
   Combo1.ListIndex = 0
   MBox1.Text = Format(Date, "DD/MM/YYYY")
End Sub




