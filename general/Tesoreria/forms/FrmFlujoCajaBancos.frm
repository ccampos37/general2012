VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Begin VB.Form FrmFlujoCajaBancos 
   Caption         =   "Form1"
   ClientHeight    =   2430
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6645
   LinkTopic       =   "Form1"
   ScaleHeight     =   2430
   ScaleWidth      =   6645
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraDetallado 
      Height          =   1590
      Left            =   120
      TabIndex        =   2
      Top             =   90
      Width           =   6345
      Begin MSComCtl2.DTPicker DTPickerFecFinal 
         Height          =   300
         Left            =   4125
         TabIndex        =   3
         Top             =   270
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   529
         _Version        =   393216
         Format          =   16187393
         CurrentDate     =   37474
      End
      Begin MSComCtl2.DTPicker DTPickerFecInicio 
         Height          =   300
         Left            =   1140
         TabIndex        =   4
         Top             =   270
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   529
         _Version        =   393216
         Format          =   16187393
         CurrentDate     =   37474
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuempresa 
         Height          =   315
         Left            =   1125
         TabIndex        =   5
         Top             =   840
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
      Begin VB.Label Lblempresa 
         AutoSize        =   -1  'True
         Caption         =   "Empresa :"
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   960
         Width           =   705
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha Final"
         Height          =   285
         Left            =   3225
         TabIndex        =   7
         Top             =   315
         Width           =   840
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Inicial"
         Height          =   300
         Left            =   75
         TabIndex        =   6
         Top             =   285
         Width           =   930
      End
   End
   Begin VB.CommandButton cmdBotones 
      Caption         =   "&Cancelar"
      Height          =   360
      Index           =   1
      Left            =   2835
      TabIndex        =   1
      Top             =   1770
      Width           =   1215
   End
   Begin VB.CommandButton cmdBotones 
      Caption         =   "&Aceptar"
      Height          =   360
      Index           =   0
      Left            =   1485
      TabIndex        =   0
      Top             =   1770
      Width           =   1215
   End
End
Attribute VB_Name = "FrmFlujoCajaBancos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim valorop As String
Dim valoroptext As String
Private Sub Form_Load()
  Dim cFecha As Date
  DTPickerFecInicio.Value = Format("01/" & Format(Month(VGParamSistem.fechatrabajo), "00") & "/" & Year(VGParamSistem.fechatrabajo), "dd/mm/yyyy")
  cFecha = Format("01/" & Format(Month(VGParamSistem.fechatrabajo) + 1, "00") & "/" & Year(VGParamSistem.fechatrabajo), "dd/mm/yyyy")
  DTPickerFecFinal.Value = Format(cFecha - 1, "dd/mm/yyyy")
  Call Ctr_Ayuempresa.Conexion(VGCNx)
End Sub

Private Sub cmdBotones_Click(index As Integer)
  Select Case index
    Case 0:
      Call Impresionflujocajabancos
    Case 1:
      Unload Me
  End Select
End Sub

Sub Impresionflujocajabancos()
Dim arrform() As Variant, arrparm() As Variant
    ReDim arrparm(4)
    ReDim arrform(1)
    arrparm(0) = VGParamSistem.BDEmpresa
    arrparm(1) = Format(DTPickerFecInicio.Value, "dd/mm/yyyy")
    arrparm(2) = Format(DTPickerFecFinal.Value, "dd/mm/yyyy")
    arrparm(3) = IIf(Ctr_Ayuempresa.xclave = Empty, "%%", Trim(Ctr_Ayuempresa.xclave))
    
    arrform(0) = "rangofecha=' DEL : " & Format(DTPickerFecInicio.Value, "dd/mm/yyyy") & " AL " & Format(DTPickerFecFinal.Value, "dd/mm/yyyy") & "'"
    
    Call ImpresionRptProc("te_flujoCajaBancos.rpt", arrform, arrparm)
End Sub

