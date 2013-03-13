VERSION 5.00
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRepGer 
   Caption         =   "Reporte gerencial"
   ClientHeight    =   3255
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4965
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3255
   ScaleWidth      =   4965
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   2055
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   4815
      Begin MSComCtl2.DTPicker DtHasta 
         Height          =   285
         Left            =   1425
         TabIndex        =   5
         Top             =   1485
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   503
         _Version        =   393216
         Format          =   16318465
         CurrentDate     =   39818
      End
      Begin MSComCtl2.DTPicker DTDesde 
         Height          =   285
         Left            =   1425
         TabIndex        =   6
         Top             =   1080
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   503
         _Version        =   393216
         Format          =   16318465
         CurrentDate     =   39818
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyudaPunto 
         Height          =   315
         Left            =   1425
         TabIndex        =   7
         Top             =   675
         Width           =   3105
         _ExtentX        =   5477
         _ExtentY        =   556
         XcodMaxLongitud =   2
         xcodwith        =   100
         NomTabla        =   "vt_puntoventa"
         TituloAyuda     =   "Ayuda de Punto de Ventas"
         ListaCampos     =   "puntovtacodigo(1),puntovtadescripcion(1)"
         XcodCampo       =   "puntovtacodigo"
         XListCampo      =   "puntovtadescripcion"
         ListaCamposDescrip=   "Codigo,Descripcion"
         ListaCamposText =   "puntovtacodigo,puntovtadescripcion"
         Requerido       =   0   'False
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuempresa 
         Height          =   315
         Left            =   1410
         TabIndex        =   11
         Top             =   240
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   556
         XcodMaxLongitud =   3
         xcodwith        =   300
         NomTabla        =   "co_multiempresas"
         TituloAyuda     =   "Busqueda de Empresas"
         ListaCampos     =   "empresacodigo(1),empresadescripcion(1),agentederetencion(1)"
         XcodCampo       =   "empresacodigo"
         XListCampo      =   "empresadescripcion"
         ListaCamposDescrip=   "Codigo,Descripcion"
         ListaCamposText =   "empresacodigo,empresadescripcion,agentederetencion"
      End
      Begin VB.Label Leempresa 
         AutoSize        =   -1  'True
         Caption         =   "Empresa :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   180
         TabIndex        =   12
         Top             =   300
         Width           =   1065
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Desde :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   120
         TabIndex        =   10
         Top             =   1170
         Width           =   825
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Hasta :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   120
         TabIndex        =   9
         Top             =   1530
         Width           =   750
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Punto Venta :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   120
         TabIndex        =   8
         Top             =   765
         Width           =   1380
      End
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   2250
      TabIndex        =   1
      Top             =   2430
      Width           =   1230
   End
   Begin VB.CommandButton CmdImp 
      Caption         =   "Imprimir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   900
      TabIndex        =   0
      Top             =   2430
      Width           =   1230
   End
   Begin ctrlayuda_f.Ctr_Ayuda Ctr_AYudaMoneda 
      Height          =   315
      Left            =   1620
      TabIndex        =   2
      Top             =   3330
      Visible         =   0   'False
      Width           =   2205
      _ExtentX        =   3889
      _ExtentY        =   556
      XcodMaxLongitud =   3
      xcodwith        =   200
      NomTabla        =   "gr_moneda"
      TituloAyuda     =   "Ayuda de Moneda"
      ListaCampos     =   "monedacodigo(1),monedadescripcion(1)"
      XcodCampo       =   "monedacodigo"
      XListCampo      =   "monedadescripcion"
      ListaCamposDescrip=   "Codigo,Descripcion"
      ListaCamposText =   "monedacodigo,monedadescripcion"
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Moneda :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   315
      TabIndex        =   3
      Top             =   3420
      Visible         =   0   'False
      Width           =   810
   End
End
Attribute VB_Name = "FrmRepGer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim formulas(1) As Variant
Dim Param(2) As Variant

Param(0) = VGParamSistem.BDEmpresa
Param(1) = VGParamSistem.TipoCambio

formulas(0) = "@tipodecambio='" & VGParamSistem.TipoCambio & "'"

Call ImpresionRptProc("vt_reportegerente.rpt", formulas, Param, , "Reporte al Gerente")

End Sub

Private Sub CmdImp_Click()
Dim formulas(4) As Variant
Dim Param(6) As Variant

If Ctr_Ayuempresa.xclave = "" Then
    MsgBox "Ingrese Codigo de Empresa", vbInformation, "Sistema"
    Ctr_Ayuempresa.SetFocus
    Exit Sub
End If

If DTDesde > DtHasta Then
    MsgBox "El fecha Inicial mayor Fecha Final", vbInformation, "Sistema"
    TxtTc.SetFocus
    Exit Sub
End If

Param(0) = VGParamSistem.BDEmpresa
Param(1) = DTDesde.Value
Param(2) = DtHasta.Value
Param(3) = IIf(Ctr_AyudaPunto.xclave <> "", Ctr_AyudaPunto.xclave, "%")
Param(4) = Ctr_Ayuempresa.xclave

formulas(0) = "fechaini='" & Format(DTDesde.Value, "dd/mm/yyyy") & "'"
formulas(1) = "fechafin='" & Format(DtHasta.Value, "dd/mm/yyyy") & "'"
If Ctr_AyudaPunto.xclave <> "" Then
    formulas(2) = "puntosvtas='" & Ctr_AyudaPunto.xnombre & "'"
 Else
     formulas(2) = "puntosvtas='TODOS'"
 End If
formulas(3) = "empresa='" & Ctr_Ayuempresa.xnombre & "'"
Call ImpresionRptProc("vt_reportegerente.rpt", formulas, Param, , "Reporte al Gerente")
End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub Ctr_AyudaPunto_KeyPress(KeyAscii As Integer)
If keyscii = 13 Then SendKeys "{tab}"
End Sub


Private Sub DtDesde_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{tab}"
End Sub


Private Sub DtHasta_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{tab}"
End Sub


Private Sub Form_Load()
Call Ctr_AyudaPunto.Conexion(VGCNx)
Call Ctr_AYudaMoneda.Conexion(VGCNx)
Call Ctr_Ayuempresa.Conexion(VGCNx)
DTDesde.Value = Date
DtHasta.Value = Date
End Sub

Private Sub TxtTc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub TxtTc_LostFocus()
If Len(Trim(TxtTc.Text)) = 0 Then TxtTc.Text = 1

End Sub


