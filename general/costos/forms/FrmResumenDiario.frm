VERSION 5.00
Object = "{4D137D9C-00A6-4458-9B46-7E95DB76D55B}#9.0#0"; "TextFer.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmresumenDiario 
   Caption         =   "Costo Diario"
   ClientHeight    =   3330
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   5685
   Icon            =   "FrmResumenDiario.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3330
   ScaleWidth      =   5685
   StartUpPosition =   3  'Windows Default
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
      Height          =   2655
      Left            =   3960
      TabIndex        =   9
      Top             =   360
      Width           =   1455
      Begin VB.CommandButton cmdImp 
         Caption         =   "&Imprimir"
         Height          =   675
         Left            =   360
         Picture         =   "FrmResumenDiario.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   480
         Width           =   775
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "&Salir"
         Height          =   675
         Left            =   360
         Picture         =   "FrmResumenDiario.frx":110C
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1680
         Width           =   780
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
      TabIndex        =   4
      Top             =   360
      Width           =   3615
      Begin VB.CheckBox Check1 
         Caption         =   "Incluye Detalle"
         Height          =   495
         Left            =   1920
         TabIndex        =   12
         Top             =   -120
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   1200
         TabIndex        =   5
         Top             =   360
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         Format          =   16711683
         CurrentDate     =   39541
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   1200
         TabIndex        =   6
         Top             =   840
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         Format          =   16711683
         CurrentDate     =   39541
      End
      Begin VB.Label Label1 
         Caption         =   "Desde"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   960
         Width           =   1215
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
      Height          =   1095
      Left            =   240
      TabIndex        =   0
      Top             =   1920
      Width           =   3615
      Begin VB.OptionButton Option01 
         Caption         =   "Soles"
         Height          =   255
         Left            =   480
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton Option02 
         Caption         =   "Dolares"
         Height          =   255
         Left            =   480
         TabIndex        =   1
         Top             =   600
         Width           =   855
      End
      Begin TextFer.TxFer TxFertipocambio 
         Height          =   375
         Left            =   1800
         TabIndex        =   3
         Top             =   360
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
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
         TipoDato        =   1
         NumeroDecimales =   4
         MarcarTextoAlEnfoque=   -1  'True
      End
   End
End
Attribute VB_Name = "FrmresumenDiario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdImp_Click()
Call imprimirresumen
End Sub
Private Sub imprimirresumen()
Dim aform(5) As Variant
Dim aparam(8) As Variant
Dim valor As Double
Dim fecha As String
aform(0) = "@mes='" & DesMes(Format(Month(DTPicker1.Value), "00")) & Str(Year(DTPicker1)) & "'"
aparam(0) = VGParametros.BaseOrigen
aparam(1) = VGCNx.DefaultDatabase
aparam(2) = DTPicker1.Value
aparam(3) = DTPicker2.Value
aparam(4) = 1
aparam(5) = TxFertipocambio.valor
aform(2) = "@valorcambio='" & aparam(5) & "'"
fecha = Format(DTPicker2.Value, "dd/mm/yyyy")
If DTPicker2.Value <> DTPicker1.Value Then
   aform(4) = "@Fecha='  DEL : " & DTPicker1.Value & " AL : " & fecha & "'"
 Else
   aform(4) = "@Fecha='  DEL DIA : " & fecha & "'"
 End If
 
If Option02 Then
   aform(1) = "@moneda='DOLARES'"
   aform(3) = "@tipocambio='02'"
   aparam(6) = "02"
 Else
   aform(1) = "@moneda='SOLES'"
   aform(3) = "@tipocambio='01'"
   aparam(6) = "01"
End If
aparam(7) = DTPicker2.Value - DTPicker1.Value + 1
Call ImpresionRptProc("cs_ResumenDiario.rpt", aform, aparam, , "Resumen Diario")
If Check1.Value Then
   aparam(4) = 1
   Call ImpresionRptProc("cs_ResumenDiarioDetallado.rpt", aform, aparam, , "Resumen Diario deatallado")
End If
End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub DTPicker2_Change()
TxFertipocambio.valor = XRecuperaTipoCambio(DTPicker2.Value, Venta, VGCNx)
TxFertipocambio.valor = IIf(TxFertipocambio.valor = 0, 1, TxFertipocambio.valor)
End Sub

Private Sub Form_Load()
Option01.Value = True
DTPicker1.Value = fecha(1, VGParamSistem.FechaTrabajo)
DTPicker2.Value = VGParamSistem.FechaTrabajo
TxFertipocambio.valor = XRecuperaTipoCambio(DTPicker2.Value, Venta, VGCNx)
TxFertipocambio.valor = IIf(TxFertipocambio.valor = 0, 1, TxFertipocambio.valor)
End Sub
Private Sub validacion()
Set VGCommandoSP = New ADODB.Command
VGCommandoSP.ActiveConnection = VGgeneral
VGCommandoSP.CommandType = adCmdStoredProc
VGCommandoSP.CommandText = "cs_actualizacostosDiarios_pro"
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

Private Sub Option01_Click()
If Option01.Value = True Then
  TxFertipocambio.Visible = False
End If
End Sub

Private Sub Option02_Click()
If Option02.Value = True Then
  TxFertipocambio.Visible = True
  TxFertipocambio.valor = XRecuperaTipoCambio(DTPicker2.Value, Venta, VGCNx)
  TxFertipocambio.SetFocus
End If
End Sub



