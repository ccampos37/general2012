VERSION 5.00
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form RptNotaabono 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nota Cargo / Abono"
   ClientHeight    =   3630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5490
   Icon            =   "RptNotaabono.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   5490
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   2205
      Left            =   180
      TabIndex        =   3
      Top             =   870
      Width           =   4965
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1230
         TabIndex        =   7
         Text            =   "Combo2"
         Top             =   1710
         Width           =   1365
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1200
         TabIndex        =   8
         Text            =   "Combo1"
         Top             =   1290
         Width           =   1425
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuda1 
         Height          =   285
         Left            =   1170
         TabIndex        =   13
         Top             =   210
         Width           =   3555
         _ExtentX        =   6271
         _ExtentY        =   503
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
         Left            =   1170
         TabIndex        =   14
         Top             =   540
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
         Left            =   3270
         TabIndex        =   15
         Top             =   540
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuda2 
         Height          =   285
         Left            =   1200
         TabIndex        =   9
         Top             =   930
         Width           =   3555
         _ExtentX        =   6271
         _ExtentY        =   503
         Enabled         =   0   'False
         XcodMaxLongitud =   0
      End
      Begin VB.Label Label2 
         Caption         =   "Ordenado"
         Height          =   255
         Index           =   2
         Left            =   300
         TabIndex        =   12
         Top             =   1710
         Width           =   885
      End
      Begin VB.Label Label2 
         Caption         =   "Estado"
         Height          =   255
         Index           =   1
         Left            =   300
         TabIndex        =   11
         Top             =   1350
         Width           =   885
      End
      Begin VB.Label Label2 
         Caption         =   "Tipo Dcmto"
         Height          =   255
         Index           =   0
         Left            =   270
         TabIndex        =   10
         Top             =   960
         Width           =   885
      End
      Begin VB.Label Label1 
         Caption         =   "Hasta"
         Height          =   225
         Index           =   2
         Left            =   2670
         TabIndex        =   6
         Top             =   600
         Width           =   915
      End
      Begin VB.Label Label1 
         Caption         =   "Desde"
         Height          =   225
         Index           =   1
         Left            =   300
         TabIndex        =   5
         Top             =   600
         Width           =   765
      End
      Begin VB.Label Label1 
         Caption         =   "Cliente"
         Height          =   225
         Index           =   0
         Left            =   300
         TabIndex        =   4
         Top             =   240
         Width           =   915
      End
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Left            =   150
      TabIndex        =   0
      Top             =   240
      Width           =   4995
      Begin VB.OptionButton Option2 
         Caption         =   "Fecha Vencimiento"
         Height          =   195
         Left            =   2550
         TabIndex        =   2
         Top             =   210
         Width           =   1725
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Fecha Emision"
         Height          =   195
         Left            =   360
         TabIndex        =   1
         Top             =   210
         Value           =   -1  'True
         Width           =   1725
      End
   End
End
Attribute VB_Name = "RptNotaabono"
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
   Combo1.AddItem "C-Cancelado"
   Combo1.AddItem "P-Pendiente"
   Combo1.ListIndex = 0
   
   Combo2.Clear
   Combo2.AddItem "C-Cliente"
   Combo2.AddItem "T-Tipo Documento"
   Combo2.ListIndex = 0
   
   MBox1.Text = Format(Date, "DD/MM/YYYY")
   MBox2.Text = Format(Date, "DD/MM/YYYY")
End Sub







