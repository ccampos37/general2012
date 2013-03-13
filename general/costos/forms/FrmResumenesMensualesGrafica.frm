VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmResumenesMensualesGrafica 
   Caption         =   "Form1"
   ClientHeight    =   4155
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   5940
   LinkTopic       =   "Form1"
   ScaleHeight     =   4155
   ScaleWidth      =   5940
   StartUpPosition =   3  'Windows Default
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
      Left            =   4080
      TabIndex        =   11
      Top             =   240
      Width           =   1215
      Begin VB.OptionButton Option02 
         Caption         =   "Dolares"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   840
         Width           =   855
      End
      Begin VB.OptionButton Option01 
         Caption         =   "Soles"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   360
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
      Left            =   360
      TabIndex        =   6
      Top             =   240
      Width           =   3615
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   1200
         TabIndex        =   7
         Top             =   960
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         Format          =   16973827
         CurrentDate     =   39541
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   1200
         TabIndex        =   8
         Top             =   360
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         Format          =   16973827
         CurrentDate     =   39541
      End
      Begin VB.Label Label2 
         Caption         =   "Mes Actual"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Desde"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1080
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
      Height          =   1215
      Left            =   1200
      TabIndex        =   3
      Top             =   2640
      Width           =   2175
      Begin VB.CommandButton CmdSalir 
         Caption         =   "&Salir"
         Height          =   795
         Left            =   1200
         Picture         =   "FrmResumenesMensualesGrafica.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   795
      End
      Begin VB.CommandButton cmdImp 
         Caption         =   "&Imprimir"
         Height          =   675
         Left            =   120
         Picture         =   "FrmResumenesMensualesGrafica.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   360
         Width           =   775
      End
   End
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
      Left            =   360
      TabIndex        =   0
      Top             =   1800
      Width           =   4935
      Begin VB.CheckBox Checkpersonal 
         Caption         =   "Personal"
         Height          =   255
         Left            =   960
         TabIndex        =   2
         Top             =   360
         Width           =   1095
      End
      Begin VB.CheckBox Checkgastos 
         Caption         =   "Gastos Generales"
         Height          =   255
         Left            =   2640
         TabIndex        =   1
         Top             =   360
         Width           =   1935
      End
   End
End
Attribute VB_Name = "FrmResumenesMensualesGrafica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdImp_Click()
Call imprimirresumen
End Sub
Private Sub imprimirresumen()
Dim aform(1) As Variant
Dim aparam(4) As Variant
Dim valor As Double
aparam(0) = VGCNx.DefaultDatabase
aparam(1) = DTPicker1.Value
aparam(2) = DTPicker2.Value
If Option02 Then
   aform(0) = "@moneda='DOLARES'"
   aparam(3) = "02"
 Else
   aparam(3) = "01"
   aform(0) = "@moneda='SOLES'"
End If
Call ImpresionRptProc("cs_costounitarioxmeses.rpt", aform, aparam, , "Costo Unitario x meses")
If Checkpersonal.Value Then Call ImpresionRptProc("cs_costounitariopersonalxmeses.rpt", aform, aparam, , "Costo unitario x meses")
If Checkgastos.Value Then Call ImpresionRptProc("cs_costounitariogastosxmeses.rpt", aform, aparam, , "Costo Unitario Detallado x meses")
End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub DTPicker2_Change()
DTPicker1.Value = Fecha(1, DTPicker2.Value - (VGParametros.mesesreferencia) * 30)
End Sub

Private Sub Form_Load()
Option01.Value = True
DTPicker2.Value = Fecha(2, VGParamSistem.FechaTrabajo)
DTPicker1.Value = Fecha(1, DTPicker2.Value - (VGParametros.mesesreferencia) * 30)
Checkgastos.Value = 1
Checkpersonal.Value = 1
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


