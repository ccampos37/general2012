VERSION 5.00
Object = "{4D137D9C-00A6-4458-9B46-7E95DB76D55B}#9.0#0"; "TextFer.ocx"
Begin VB.Form frmRepComprob 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reporte de Comprobantes"
   ClientHeight    =   2340
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4020
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2340
   ScaleWidth      =   4020
   Begin VB.CommandButton Command2 
      Caption         =   "&Aceptar"
      Height          =   360
      Left            =   2040
      TabIndex        =   2
      Top             =   1905
      Width           =   1140
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   360
      Left            =   840
      TabIndex        =   1
      Top             =   1905
      Width           =   1140
   End
   Begin VB.Frame Frame1 
      Caption         =   "Rango de Comprobantes"
      Height          =   1710
      Left            =   90
      TabIndex        =   0
      Top             =   75
      Width           =   3855
      Begin TextFer.TxFer TxCompIni 
         Height          =   330
         Left            =   165
         TabIndex        =   4
         Top             =   495
         Width           =   3420
         _ExtentX        =   6033
         _ExtentY        =   582
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
         ColorIlumina    =   -2147483624
         SaltarAlEnter   =   -1  'True
         Valor           =   ""
         NoCaracteres    =   "0123456789"
         MarcarTextoAlEnfoque=   -1  'True
         NoRangoCadena   =   -1  'True
      End
      Begin TextFer.TxFer Txcomprfin 
         Height          =   330
         Left            =   165
         TabIndex        =   5
         Top             =   1140
         Width           =   3420
         _ExtentX        =   6033
         _ExtentY        =   582
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
         ColorIlumina    =   -2147483624
         SaltarAlEnter   =   -1  'True
         Valor           =   ""
         NoCaracteres    =   "0123456789"
         MarcarTextoAlEnfoque=   -1  'True
         NoRangoCadena   =   -1  'True
      End
      Begin VB.Label Label2 
         Caption         =   "Comprobante Final :"
         Height          =   225
         Left            =   180
         TabIndex        =   6
         Top             =   915
         Width           =   2190
      End
      Begin VB.Label Label1 
         Caption         =   "Comprobante Inicial :"
         Height          =   225
         Left            =   180
         TabIndex        =   3
         Top             =   255
         Width           =   2190
      End
   End
End
Attribute VB_Name = "frmRepComprob"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Private Sub Command1_Click()
    Call imprimir
End Sub

Private Sub Form_Load()
    Width = 4110
    heigth = 2715
End Sub
Private Sub imprimir()
'FIXIT: Declare 'arrform' con un tipo de datos de enlace en tiempo de compilación          FixIT90210ae-R1672-R1B8ZE
Dim arrform(0) As Variant, arrparm(7) As Variant
    Screen.MousePointer = 11
    arrparm(0) = Trim$(VGParamSistem.BDEmpresa)
    arrparm(1) = VGParametros.empresacodigo
    arrparm(2) = Trim$(VGParamSistem.Anoproceso)
    arrparm(3) = Trim$(VGParamSistem.Mesproceso)
    arrparm(4) = Trim$(TxCompIni.Text)
    arrparm(5) = Trim$(Txcomprfin.Text)
    arrparm(6) = Trim$(SubAsiento)
    Call ImpresionRptProc("rptVoucherComprobRang.rpt", arrform, arrparm)
    Screen.MousePointer = 1
End Sub
