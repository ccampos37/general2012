VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Begin VB.Form FrmResumenGastosTesoreria 
   Caption         =   "Resumenes de Tesoreria"
   ClientHeight    =   5535
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6960
   LinkTopic       =   "Form1"
   ScaleHeight     =   5535
   ScaleWidth      =   6960
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBotones 
      Caption         =   "&Aceptar"
      Height          =   360
      Index           =   0
      Left            =   1770
      TabIndex        =   20
      Top             =   4890
      Width           =   1215
   End
   Begin VB.CommandButton cmdBotones 
      Caption         =   "&Cancelar"
      Height          =   360
      Index           =   1
      Left            =   3120
      TabIndex        =   19
      Top             =   4890
      Width           =   1215
   End
   Begin VB.Frame fraDetallado 
      Height          =   3990
      Left            =   285
      TabIndex        =   3
      Top             =   690
      Width           =   6345
      Begin VB.ComboBox cmbtipmov 
         Height          =   315
         ItemData        =   "FrmResumenGastosTesoreria.frx":0000
         Left            =   1110
         List            =   "FrmResumenGastosTesoreria.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1890
         Width           =   1770
      End
      Begin VB.Frame Frame2 
         Height          =   720
         Left            =   135
         TabIndex        =   4
         Top             =   2970
         Visible         =   0   'False
         Width           =   5925
         Begin VB.OptionButton Opt2 
            Caption         =   "Solo Transferencias"
            Height          =   195
            Index           =   0
            Left            =   135
            TabIndex        =   7
            Top             =   300
            Width           =   1740
         End
         Begin VB.OptionButton Opt2 
            Caption         =   "Sin Transferencias"
            Height          =   195
            Index           =   1
            Left            =   2040
            TabIndex        =   6
            Top             =   330
            Width           =   1740
         End
         Begin VB.OptionButton Opt2 
            Caption         =   "Todas"
            Height          =   195
            Index           =   2
            Left            =   4110
            TabIndex        =   5
            Top             =   360
            Width           =   1740
         End
      End
      Begin MSComCtl2.DTPicker DTPickerFecFinal 
         Height          =   300
         Left            =   4125
         TabIndex        =   9
         Top             =   270
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   529
         _Version        =   393216
         Format          =   57802753
         CurrentDate     =   37474
      End
      Begin MSComCtl2.DTPicker DTPickerFecInicio 
         Height          =   300
         Left            =   1140
         TabIndex        =   10
         Top             =   270
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   529
         _Version        =   393216
         Format          =   57802753
         CurrentDate     =   37474
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_concepto1 
         Height          =   315
         Left            =   1125
         TabIndex        =   11
         Top             =   1410
         Width           =   4845
         _ExtentX        =   8546
         _ExtentY        =   556
         XcodMaxLongitud =   4
         xcodwith        =   800
         NomTabla        =   "te_conceptocaja"
         TituloAyuda     =   "Ayuda de Conceptos"
         ListaCampos     =   "conceptocodigo(1),conceptodescripcion(1)"
         XcodCampo       =   "conceptocodigo"
         XListCampo      =   "conceptodescripcion"
         ListaCamposDescrip=   "Código,Descripción"
         ListaCamposText =   "conceptocodigo,conceptodescripcion"
         Requerido       =   0   'False
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuempresa 
         Height          =   315
         Left            =   1125
         TabIndex        =   12
         Top             =   2400
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   556
         XcodMaxLongitud =   3
         xcodwith        =   300
         NomTabla        =   "co_multiempresas"
         TituloAyuda     =   "Busqueda de Empresas"
         ListaCampos     =   "empresacodigo(1),empresadescripcion(1)"
         XcodCampo       =   "empresacodigo"
         XListCampo      =   "empresadescripcion"
         ListaCamposDescrip=   "Codigo,Descripcion"
         ListaCamposText =   "empresacodigo,empresadescripcion"
         Requerido       =   0   'False
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyudaCaja 
         Height          =   315
         Left            =   1080
         TabIndex        =   22
         Top             =   960
         Width           =   4260
         _ExtentX        =   7514
         _ExtentY        =   556
         XcodMaxLongitud =   11
         xcodwith        =   400
         NomTabla        =   "te_codigocaja"
         TituloAyuda     =   "Busqueda de Caja"
         ListaCampos     =   "cajacodigo(1),cajadescripcion(1)"
         XcodCampo       =   "cajacodigo"
         XListCampo      =   "cajadescripcion"
         ListaCamposDescrip=   "Codigo,Descripcion"
         ListaCamposText =   "cajacodigo,cajadescripcion"
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Inicial"
         Height          =   300
         Left            =   75
         TabIndex        =   18
         Top             =   285
         Width           =   930
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha Final"
         Height          =   285
         Left            =   3225
         TabIndex        =   17
         Top             =   315
         Width           =   840
      End
      Begin VB.Label lmon 
         Caption         =   "Caja"
         Height          =   225
         Left            =   150
         TabIndex        =   16
         Top             =   1005
         Width           =   885
      End
      Begin VB.Label Label7 
         Caption         =   "Concepto"
         Height          =   285
         Left            =   150
         TabIndex        =   15
         Top             =   1470
         Width           =   885
      End
      Begin VB.Label Label8 
         Caption         =   "Tipo de Movimiento"
         Height          =   570
         Left            =   135
         TabIndex        =   14
         Top             =   1845
         Width           =   885
      End
      Begin VB.Label Lblempresa 
         AutoSize        =   -1  'True
         Caption         =   "Empresa :"
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   2460
         Width           =   705
      End
   End
   Begin VB.Frame Frame1 
      Height          =   675
      Left            =   285
      TabIndex        =   0
      Top             =   0
      Width           =   6330
      Begin VB.OptionButton Opt 
         Caption         =   "Todo"
         Height          =   300
         Index           =   2
         Left            =   4080
         TabIndex        =   21
         Top             =   240
         Width           =   1440
      End
      Begin VB.OptionButton Opt 
         Caption         =   "Caja"
         Height          =   300
         Index           =   0
         Left            =   150
         TabIndex        =   2
         Top             =   270
         Width           =   1440
      End
      Begin VB.OptionButton Opt 
         Caption         =   "Banco"
         Height          =   300
         Index           =   1
         Left            =   2160
         TabIndex        =   1
         Top             =   240
         Width           =   1440
      End
   End
End
Attribute VB_Name = "FrmResumenGastosTesoreria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim valorop As String
Dim valoroptext As String
Private Sub Form_Load()
  Dim cFecha As Date
  Opt(2).Value = True
  Me.Width = 6860
  Me.Height = 6795
  DTPickerFecInicio.Value = Format("01/" & Format(Month(VGParamSistem.fechatrabajo), "00") & "/" & Year(VGParamSistem.fechatrabajo), "dd/mm/yyyy")
  cFecha = Format("01/" & Format(Month(VGParamSistem.fechatrabajo) + 1, "00") & "/" & Year(VGParamSistem.fechatrabajo), "dd/mm/yyyy")
  DTPickerFecFinal.Value = Format(cFecha - 1, "dd/mm/yyyy")
  Call Ctr_concepto1.Conexion(VGcnx)
  Call Ctr_Ayuempresa.Conexion(VGcnx)
  Call Ctr_AyudaCaja.Conexion(VGcnx)
  cmbtipmov.ListIndex = 2
  opt2(2).Value = True
End Sub

Private Sub cmdBotones_Click(index As Integer)
  Select Case index
    Case 0:
      Call ImpresionEstadoCtaCteResumen
    Case 1:
      Unload Me
  End Select
End Sub

Sub ImpresionEstadoCtaCteResumen()
Dim arrform() As Variant, arrparm() As Variant
    ReDim arrparm(7)
    ReDim arrform(1)
    arrparm(0) = VGParamSistem.BDEmpresa
    arrparm(1) = IIf(Opt(2).Value = True, "X", IIf(Opt(0).Value = True, "C", "B"))
    arrparm(2) = IIf(Ctr_AyudaCaja.xclave = Empty, "%%", Trim(Ctr_AyudaCaja.xclave))
    arrparm(3) = Format(DTPickerFecInicio.Value, "dd/mm/yyyy")
    arrparm(4) = Format(DTPickerFecFinal.Value, "dd/mm/yyyy")
    arrparm(5) = IIf(Trim(Ctr_concepto1.xclave) = "", "%%", Trim(Ctr_concepto1.xclave))
'    arrparm(6) = Left(Trim(cmbtipmov.Text), 2)
'    arrparm(7) = valorop
    arrparm(6) = IIf(Ctr_Ayuempresa.xclave = Empty, "%%", Trim(Ctr_Ayuempresa.xclave))
    
    arrform(0) = "rangofecha=' DEL : " & Format(DTPickerFecInicio.Value, "dd/mm/yyyy") & " AL " & Format(DTPickerFecFinal.Value, "dd/mm/yyyy") & "'"
    
    Call ImpresionRptProc("te_ResumenesdeCaja.rpt", arrform, arrparm)
End Sub

Sub ConfiguraCajaBanco(valor As Boolean)
    lmon.Visible = valor
  
End Sub


Private Sub Opt_Click(index As Integer)
  Select Case index
    Case 0:
       Call ConfiguraCajaBanco(True)
    
    Case 1:
       Call ConfiguraCajaBanco(False)
  End Select

End Sub

Private Sub Opt2_Click(index As Integer)
    Select Case index
        Case 0:
            valorop = "1"
        Case 1:
            valorop = "0"
        Case 2:
            valorop = "%%"
    End Select
    valoroptext = opt2(index).Caption
End Sub

