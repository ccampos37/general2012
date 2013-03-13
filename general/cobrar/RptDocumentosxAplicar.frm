VERSION 5.00
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "CONTROLAYUDA.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form RptDocumentosxAplicar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Documentos por Aplicar"
   ClientHeight    =   4275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5850
   Icon            =   "RptDocumentosxAplicar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   5850
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Height          =   3195
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   5445
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1770
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   2730
         Width           =   1275
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1770
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   2340
         Width           =   1275
      End
      Begin VB.Frame Frame1 
         Height          =   705
         Left            =   150
         TabIndex        =   7
         Top             =   180
         Width           =   2535
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
            Left            =   1290
            TabIndex        =   8
            Top             =   210
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label1 
            Caption         =   "Hasta Fecha"
            Height          =   255
            Index           =   1
            Left            =   150
            TabIndex        =   9
            Top             =   240
            Width           =   1245
         End
      End
      Begin VB.Frame Frame2 
         Height          =   1335
         Left            =   150
         TabIndex        =   1
         Top             =   900
         Width           =   5145
         Begin VB.OptionButton Option1 
            Caption         =   "Relacion x Banco"
            Height          =   195
            Index           =   2
            Left            =   180
            TabIndex        =   4
            Top             =   870
            Width           =   1935
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Todos Movimientos"
            Height          =   195
            Index           =   0
            Left            =   180
            TabIndex        =   3
            Top             =   270
            Width           =   1695
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Relacion x Vendedor"
            Height          =   195
            Index           =   1
            Left            =   180
            TabIndex        =   2
            Top             =   570
            Width           =   1935
         End
         Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuda1 
            Height          =   315
            Left            =   2160
            TabIndex        =   5
            Top             =   540
            Width           =   2925
            _ExtentX        =   5159
            _ExtentY        =   556
            Enabled         =   0   'False
            XcodMaxLongitud =   0
         End
         Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuda2 
            Height          =   315
            Left            =   2160
            TabIndex        =   6
            Top             =   900
            Width           =   2925
            _ExtentX        =   5159
            _ExtentY        =   556
            Enabled         =   0   'False
            XcodMaxLongitud =   0
         End
      End
      Begin VB.Label Label2 
         Caption         =   "Hoja Resumen"
         Height          =   165
         Index           =   1
         Left            =   390
         TabIndex        =   13
         Top             =   2730
         Width           =   1125
      End
      Begin VB.Label Label2 
         Caption         =   "Resumen"
         Height          =   165
         Index           =   0
         Left            =   390
         TabIndex        =   12
         Top             =   2400
         Width           =   1125
      End
   End
End
Attribute VB_Name = "RptDocumentosxAplicar"
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
   Call CargarTipo(Combo2, 3)
   MBox1.Text = Format(Date, "DD/MM/YYYY")
End Sub








