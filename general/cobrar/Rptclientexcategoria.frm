VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Rptclientexcategoria 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cliente por Categoria"
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4470
   Icon            =   "Rptclientexcategoria.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   4470
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Height          =   1695
      Left            =   150
      TabIndex        =   0
      Top             =   150
      Width           =   4095
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1260
         TabIndex        =   6
         Text            =   "Combo2"
         Top             =   720
         Width           =   2445
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1260
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1110
         Width           =   2445
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
         TabIndex        =   2
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
         TabIndex        =   5
         Top             =   270
         Width           =   1365
      End
      Begin VB.Label Label2 
         Caption         =   "Ordenado"
         Height          =   165
         Index           =   0
         Left            =   270
         TabIndex        =   4
         Top             =   1170
         Width           =   1125
      End
      Begin VB.Label Label2 
         Caption         =   "Categoria"
         Height          =   165
         Index           =   2
         Left            =   270
         TabIndex        =   3
         Top             =   720
         Width           =   1125
      End
   End
End
Attribute VB_Name = "Rptclientexcategoria"
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
   
   Call CargarTipo(Combo2, 3)
   
   MBox1.Text = Format(Date, "DD/MM/YYYY")
End Sub





