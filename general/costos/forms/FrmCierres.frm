VERSION 5.00
Object = "{4D137D9C-00A6-4458-9B46-7E95DB76D55B}#9.0#0"; "TextFer.ocx"
Begin VB.Form FrmCierres 
   Caption         =   "Control de cierres"
   ClientHeight    =   3510
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5325
   LinkTopic       =   "Form1"
   ScaleHeight     =   3510
   ScaleWidth      =   5325
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      Height          =   975
      Left            =   2880
      TabIndex        =   6
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   975
      Left            =   840
      TabIndex        =   5
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   3855
      Begin TextFer.TxFer TxFeranno 
         Height          =   375
         Left            =   1920
         TabIndex        =   3
         Top             =   360
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         Object.CausesValidation=   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   ""
         SaltarAlEnter   =   -1  'True
         Valor           =   ""
         MarcarTextoAlEnfoque=   -1  'True
      End
      Begin TextFer.TxFer TxFermes 
         Height          =   375
         Left            =   1920
         TabIndex        =   4
         Top             =   840
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         Object.CausesValidation=   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   ""
         Valor           =   ""
         MarcarTextoAlEnfoque=   -1  'True
      End
      Begin VB.Label Label2 
         Caption         =   "Mes de cierre"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Ano de Cierre"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1575
      End
   End
End
Attribute VB_Name = "FrmCierres"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim rrsql As New ADODB.Recordset
Set rrsql = VGCNx.Execute(" update cs_sistema set mesdecierre='" & TxFeranno.valor & TxFermes.valor & "'")
Set rrsql = Nothing
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
TxFeranno.valor = Left(VGParametros.mesdecierre, 4)
TxFermes.valor = Right(VGParametros.mesdecierre, 2)
End Sub
