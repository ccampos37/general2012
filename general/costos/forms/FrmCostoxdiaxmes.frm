VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmCostoxdiaxmes 
   Caption         =   "Costo Unitario x dia del mes"
   ClientHeight    =   3210
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   5655
   LinkTopic       =   "Form1"
   ScaleHeight     =   3210
   ScaleWidth      =   5655
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "Incluye Detalle"
      Height          =   495
      Left            =   1920
      TabIndex        =   11
      Top             =   0
      Width           =   1455
   End
   Begin VB.Frame Frame3 
      Caption         =   "Moneda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   735
      Left            =   0
      TabIndex        =   8
      Top             =   1680
      Width           =   3615
      Begin VB.OptionButton Option02 
         Caption         =   "Dolares"
         Height          =   255
         Left            =   1800
         TabIndex        =   10
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton Option01 
         Caption         =   "Soles"
         Height          =   255
         Left            =   480
         TabIndex        =   9
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Rango de fechas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   1455
      Left            =   0
      TabIndex        =   3
      Top             =   120
      Width           =   3615
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   1200
         TabIndex        =   4
         Top             =   360
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         Format          =   16777219
         CurrentDate     =   39541
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   1200
         TabIndex        =   5
         Top             =   840
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         Format          =   16777219
         CurrentDate     =   39541
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Desde"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Opciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   3720
      TabIndex        =   0
      Top             =   120
      Width           =   1455
      Begin VB.CommandButton CmdSalir 
         Caption         =   "&Salir"
         Height          =   675
         Left            =   360
         Picture         =   "FrmCostoxdiaxmes.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1320
         Width           =   780
      End
      Begin VB.CommandButton cmdImp 
         Caption         =   "&Imprimir"
         Height          =   675
         Left            =   480
         Picture         =   "FrmCostoxdiaxmes.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   480
         Width           =   775
      End
   End
End
Attribute VB_Name = "FrmCostoxdiaxmes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cierre As String
Dim estado As Double
Private Sub cmdImp_Click()
Call imprimirresumen
End Sub
Private Sub imprimirresumen()
Dim aform(3) As Variant
Dim aparam(6) As Variant
Dim valor As Double
Dim fecha As String
aform(0) = "@mes='" & DesMes(Format(Month(DTPicker1.Value), "00")) & Str(Year(DTPicker1)) & "'"
aparam(0) = VGParametros.BaseOrigen
aparam(1) = VGCNx.DefaultDatabase
aparam(2) = DTPicker1.Value
aparam(3) = DTPicker2.Value
fecha = Format(DTPicker2.Value, "dd/mm/yyyy")
 
If Option02 Then
   aform(1) = "@moneda='DOLARES'"
   aform(2) = "@mon'02'"
   aparam(4) = "02"
 Else
   aform(1) = "@moneda='SOLES'"
   aform(2) = "@mon='01'"
   aparam(4) = "01"
End If
aparam(5) = DTPicker2.Value - DTPicker1.Value
Call ImpresionRptProc("cs_ResumenDiarioxmes.rpt", aform, aparam, , "Resumen Diario")
End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub


Private Sub Form_Load()
Option01.Value = True
DTPicker1.Value = fecha(1, VGParamSistem.FechaTrabajo)
DTPicker2.Value = VGParamSistem.FechaTrabajo

End Sub
