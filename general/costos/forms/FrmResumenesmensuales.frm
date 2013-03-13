VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmresumenesMensuales 
   Caption         =   "Resumenes mensuales"
   ClientHeight    =   4320
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   5805
   LinkTopic       =   "Form1"
   ScaleHeight     =   4320
   ScaleWidth      =   5805
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Caption         =   "Listar"
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
      Left            =   240
      TabIndex        =   12
      Top             =   1920
      Width           =   4935
      Begin VB.CheckBox CheckIngresos 
         Caption         =   "Produccion"
         Height          =   315
         Left            =   3600
         TabIndex        =   16
         Top             =   360
         Width           =   1215
      End
      Begin VB.CheckBox Checkalmacenes 
         Caption         =   "Almacenes"
         Height          =   255
         Left            =   2400
         TabIndex        =   15
         Top             =   360
         Width           =   1095
      End
      Begin VB.CheckBox Checkgastos 
         Caption         =   "Gastos"
         Height          =   255
         Left            =   1320
         TabIndex        =   14
         Top             =   360
         Width           =   975
      End
      Begin VB.CheckBox Checkpersonal 
         Caption         =   "Personal"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   975
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
      Height          =   1215
      Left            =   1080
      TabIndex        =   9
      Top             =   2760
      Width           =   2175
      Begin VB.CommandButton cmdImp 
         Caption         =   "&Imprimir"
         Height          =   675
         Left            =   120
         Picture         =   "FrmResumenesmensuales.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   360
         Width           =   775
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "&Salir"
         Height          =   795
         Left            =   1200
         Picture         =   "FrmResumenesmensuales.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   240
         Width           =   795
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
      Left            =   240
      TabIndex        =   4
      Top             =   360
      Width           =   3615
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   1200
         TabIndex        =   5
         Top             =   960
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         Format          =   16842755
         CurrentDate     =   39541
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   1200
         TabIndex        =   6
         Top             =   360
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         Format          =   16842755
         CurrentDate     =   39541
      End
      Begin VB.Label Label1 
         Caption         =   "Desde"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Mes Actual"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   480
         Width           =   855
      End
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
      Height          =   1455
      Left            =   3960
      TabIndex        =   1
      Top             =   360
      Width           =   1215
      Begin VB.OptionButton Option01 
         Caption         =   "Soles"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton Option02 
         Caption         =   "Dolares"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   855
      End
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Incluye Detalle"
      Height          =   495
      Left            =   2160
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "FrmresumenesMensuales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdImp_Click()
Call imprimirresumen
End Sub
Private Sub imprimirresumen()
Dim aform(1) As Variant
Dim aparam(5) As Variant
Dim valor As Double
aparam(0) = VGParametros.BaseOrigen
aparam(1) = VGCNx.DefaultDatabase
aparam(2) = DTPicker1.Value
aparam(3) = DTPicker2.Value
If Option02 Then
   aform(0) = "@moneda='DOLARES'"
   aparam(4) = "02"
 Else
   aparam(4) = "01"
   aform(0) = "@moneda='SOLES'"
End If
If Checkpersonal.Value Then Call ImpresionRptProc("cs_ResumenesMensualesPersonal.rpt", aform, aparam, , "Resumenes Mensuales Personal")
If Checkgastos.Value Then Call ImpresionRptProc("cs_ResumenesMensualesGastos.rpt", aform, aparam, , "Resumenes Mensuales Personal")
If Checkalmacenes.Value Then Call ImpresionRptProc("cs_ResumenesMensualesAlmacenes.rpt", aform, aparam, , "Resumenes Mensuales Personal")
If CheckIngresos.Value Then Call ImpresionRptProc("cs_ResumenesMensualesIngresos.rpt", aform, aparam, , "Resumenes Mensuales Personal")
End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub DTPicker2_Change()
DTPicker1.Value = fecha(1, DTPicker2.Value - (VGParametros.mesesreferencia) * 30)
End Sub

Private Sub Form_Load()
Option01.Value = True
DTPicker2.Value = fecha(2, VGParamSistem.FechaTrabajo)
DTPicker1.Value = fecha(1, DTPicker2.Value - (VGParametros.mesesreferencia) * 30)
End Sub
Private Sub validacion()
Set VGCommandoSP = New ADODB.Command
VGCommandoSP.ActiveConnection = VGgeneral
VGCommandoSP.CommandType = adCmdStoredProc
VGCommandoSP.CommandText = "cs_actualizacostos_pro"
    VGCommandoSP.Parameters.Refresh
    With VGCommandoSP
        .Parameters("@Baseorigen") = VGParametros.BaseOrigen
        .Parameters("@Basedestino") = VGCNx.DefaultDatabase
        .Parameters("@fechaini") = DTPicker1.Value
        .Parameters("@fechafin") = DTPicker2.Value
        .Parameters("@tipo") = 1
        .Parameters("@tipocambio") = TxFertipocambio.valor
        .Execute
    End With
End Sub

