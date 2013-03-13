VERSION 5.00
Object = "{4D137D9C-00A6-4458-9B46-7E95DB76D55B}#9.0#0"; "TextFer.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmResumenGeneral 
   Caption         =   "Impresion de Hoja Mensual Resumen"
   ClientHeight    =   4200
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   ScaleHeight     =   4200
   ScaleWidth      =   6030
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check2 
      Caption         =   "Incluye x Producto"
      Height          =   495
      Left            =   3840
      TabIndex        =   16
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Frame Frame4 
      Caption         =   "Tipo"
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
      Height          =   855
      Left            =   120
      TabIndex        =   13
      Top             =   240
      Width           =   5295
      Begin VB.OptionButton Option1 
         Caption         =   "Rango de meses"
         Height          =   495
         Left            =   2880
         TabIndex        =   15
         Top             =   240
         Width           =   1815
      End
      Begin VB.OptionButton OptionMes 
         Caption         =   "Mes"
         Height          =   495
         Left            =   840
         TabIndex        =   14
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Incluye Detalle"
      Height          =   495
      Left            =   1920
      TabIndex        =   12
      Top             =   1200
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
      Height          =   1095
      Left            =   120
      TabIndex        =   8
      Top             =   2880
      Width           =   3615
      Begin VB.OptionButton Option02 
         Caption         =   "Dolares"
         Height          =   255
         Left            =   480
         TabIndex        =   10
         Top             =   600
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
      Begin TextFer.TxFer TxFertipocambio 
         Height          =   375
         Left            =   1800
         TabIndex        =   11
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
      Left            =   120
      TabIndex        =   3
      Top             =   1320
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
         Format          =   16842755
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
         Format          =   16842755
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
      Height          =   2055
      Left            =   3840
      TabIndex        =   0
      Top             =   1920
      Width           =   1695
      Begin VB.CommandButton CmdSalir 
         Caption         =   "&Salir"
         Height          =   675
         Left            =   480
         Picture         =   "FrmResumenGeneral.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1200
         Width           =   780
      End
      Begin VB.CommandButton cmdImp 
         Caption         =   "&Imprimir"
         Height          =   675
         Left            =   480
         Picture         =   "FrmResumenGeneral.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   480
         Width           =   775
      End
   End
End
Attribute VB_Name = "FrmResumenGeneral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim estado As Integer
Private Sub cmdImp_Click()
Dim cierre1 As String
Dim mes As Double
estado = 0
mes = Val(Right(VGParametros.mesdecierre, 2))
If mes + 1 = 13 Then
  cierre1 = Format(Val(Left(VGParametros.mesdecierre, 4)) + 1, "0000") + "01"
Else
  cierre1 = Format(Val(Left(VGParametros.mesdecierre, 4)), "0000") + Format(mes + 1, "00")
End If
 cierre = Format(Year(DTPicker2), "0000") + Format(Month(DTPicker2), "00")

estado = 0
If OptionMes.Value = True Then
   If Month(DTPicker1.Value) <> Month(DTPicker2.Value) Then
      MsgBox (" La fechas deben ser de un solo mes ")
    Else
       If cierre <= VGParametros.mesdecierre Then
            estado = 1
            MsgBox (" Mes a Iprimir esta Cerrado , Solamente se Imprime datos Historicos")
 '           Call imprimirresumen(2)
        ElseIf cierre = cierre1 Then
               Call validacion
 '             Call imprimirresumen(2)
          Else
            MsgBox (" Mes anterior a Imprimir no esta Cerrado, Se procedea actulizar posble datos errados para rango de Meses")
              Call validacion
  '            Call imprimirresumen(2)
      End If
      estado = 1
      Call imprimirresumen(4)
    End If
  Else
    Call DTPicker2_Change
    estado = 1
    Call imprimirresumen(4)
End If
End Sub
Private Sub imprimirresumen(tipo As Integer)
Dim aform(4) As Variant
Dim aparam(7) As Variant
Dim valor As Double
If OptionMes.Value = False Then
   aform(0) = "@mes='" & DesMes(Format(Month(DTPicker1.Value), "00")) & ""
   aform(0) = aform(0) + " A " + DesMes(Format(Month(DTPicker2.Value), "00")) & Str(Year(DTPicker2)) & "'"
Else
aform(0) = "@mes='" & DesMes(Format(Month(DTPicker1.Value), "00")) & Str(Year(DTPicker1)) & "'"
End If
aparam(0) = VGParametros.BaseOrigen
aparam(1) = VGCNx.DefaultDatabase
aparam(2) = DTPicker1.Value
aparam(3) = DTPicker2.Value
aparam(4) = tipo
aparam(5) = TxFertipocambio.valor
aform(2) = "@valorcambio='" & aparam(5) & "'"
If Option02 Then
   aform(1) = "@moneda='DOLARES'"
   aform(3) = "@tipocambio='02'"
   aparam(6) = "02"
 Else
   aform(1) = "@moneda='SOLES'"
   aform(3) = "@tipocambio='01'"
   aparam(6) = "01"
End If
Call ImpresionRptProc("cs_ResumenMensual.rpt", aform, aparam, , "Resumen Mensual")
If Check1.Value Then
   If OptionMes.Value = True Then
      If estado = 1 Then
         aparam(4) = 5
         Call ImpresionRptProc("cs_ResumenMensualdetallado.rpt", aform, aparam, , "Resumen Mensual deatallado")
      End If
    Else
      aparam(4) = 5
      Call ImpresionRptProc("cs_ResumenMensualdetallado.rpt", aform, aparam, , "Resumen Mensual deatallado")
    End If
End If
End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub DTPicker2_Change()
TxFertipocambio.valor = XRecuperaTipoCambio(DTPicker2.Value, Venta, VGCNx)
End Sub

Private Sub Form_Load()
Option01.Value = True
OptionMes.Value = True
DTPicker1.Value = fecha(1, VGParamSistem.FechaTrabajo)
DTPicker2.Value = fecha(2, VGParamSistem.FechaTrabajo)
TxFertipocambio.valor = XRecuperaTipoCambio(DTPicker2.Value, Venta, VGCNx)
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


