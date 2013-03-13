VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Begin VB.Form frmRepChequesEmitidos 
   Caption         =   "Reporte de Cheques Emitidos"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6075
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   6075
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Height          =   1215
      Left            =   3360
      TabIndex        =   8
      Top             =   1680
      Width           =   2535
      Begin VB.CommandButton Cmdbotones 
         Caption         =   "&Salir"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Index           =   1
         Left            =   1440
         Picture         =   "frmRepChequesEmitidos.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton Cmdbotones 
         Caption         =   "&Aceptar"
         Height          =   675
         Index           =   0
         Left            =   360
         Picture         =   "frmRepChequesEmitidos.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Imprimir "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1215
      Left            =   240
      TabIndex        =   7
      Top             =   1680
      Width           =   2055
      Begin VB.OptionButton Optresumido 
         Caption         =   "Resumido"
         Height          =   255
         Left            =   360
         TabIndex        =   12
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton Optdetallado 
         Caption         =   "Detallado"
         Height          =   255
         Left            =   360
         TabIndex        =   11
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5655
      Begin MSComCtl2.DTPicker DTPickerFecFinal 
         Height          =   300
         Left            =   3735
         TabIndex        =   1
         Top             =   360
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   529
         _Version        =   393216
         Format          =   16580609
         CurrentDate     =   37474
      End
      Begin MSComCtl2.DTPicker DTPickerFecInicio 
         Height          =   300
         Left            =   1125
         TabIndex        =   2
         Top             =   360
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   529
         _Version        =   393216
         Format          =   16580609
         CurrentDate     =   37474
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyuDocumento 
         Height          =   435
         Left            =   1080
         TabIndex        =   5
         Top             =   915
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   767
         XcodMaxLongitud =   2
         xcodwith        =   200
         NomTabla        =   "cp_tipodocumento"
         TituloAyuda     =   "Busqueda de Tipo de  Documento"
         ListaCampos     =   "tdocumentocodigo(1),tdocumentodescripcion(1),tdocumentotipo(1),documentoretencion(1)"
         XcodCampo       =   "tdocumentocodigo"
         XListCampo      =   "tdocumentodescripcion"
         ListaCamposDescrip=   "Código,Descripción,CargoAbono,Retencion"
         ListaCamposText =   "tdocumentocodigo,tdocumentodescripcion,tdocumentotipo,documentoretencion"
         Requerido       =   0   'False
      End
      Begin VB.Label Label1 
         Caption         =   "Documento"
         Height          =   420
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   915
         Width           =   1050
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha Final"
         Height          =   285
         Left            =   2835
         TabIndex        =   4
         Top             =   405
         Width           =   840
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Inicial"
         Height          =   300
         Index           =   0
         Left            =   165
         TabIndex        =   3
         Top             =   405
         Width           =   930
      End
   End
End
Attribute VB_Name = "frmRepChequesEmitidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
  Dim cFecha As Date
  DTPickerFecInicio.Value = Fecha(1, VGParamSistem.fechatrabajo)
  Call Ctr_AyuDocumento.Conexion(VGCNx): Ctr_AyuDocumento.Filtro = "tdocumentotipo='A' and tdocumentovalidabanco=1 and  tdocumentoingplan=0"
  DTPickerFecFinal.Value = Fecha(2, VGParamSistem.fechatrabajo)
  OptDetallado.Value = True
End Sub

Private Sub cmdBotones_Click(index As Integer)
  Select Case index
    Case 0
       Call ImpresionChequesEmitidos
    Case 1
       Unload Me
  End Select
   
End Sub

Sub ImpresionChequesEmitidos()
Dim arrform() As Variant, arrparm() As Variant
Dim dato As String
ReDim arrparm(5)
ReDim arrform(1)

    arrparm(0) = VGParamSistem.BDEmpresa
    arrparm(1) = IIf(Ctr_AyuDocumento.xclave = "", "%%", Ctr_AyuDocumento.xclave)
    arrparm(2) = Format(DTPickerFecInicio.Value, "dd/mm/yyyy")
    arrparm(3) = Format(DTPickerFecFinal.Value, "dd/mm/yyyy")
    arrparm(4) = IIf(OptResumido.Value = True, "1", "0")
    dato = " xx"
    arrform(0) = "dato='" & dato & "'"
 '  arrform(0) = "xx"
    If arrparm(4) = "0" Then
       Call ImpresionRptProc("te_ListadoCheques.rpt", arrform, arrparm)
     Else
       Call ImpresionRptProc("te_ListadoChequesResumen.rpt", arrform, arrparm)
  End If
End Sub

