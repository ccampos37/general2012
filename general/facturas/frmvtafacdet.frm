VERSION 5.00
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmvtafacdet 
   Caption         =   "Ventas por factura detallado"
   ClientHeight    =   3525
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5325
   LinkTopic       =   "Form1"
   ScaleHeight     =   3525
   ScaleWidth      =   5325
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TxtTc 
      Height          =   375
      Left            =   2115
      TabIndex        =   8
      Top             =   1035
      Width           =   1590
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1230
      TabIndex        =   4
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2790
      TabIndex        =   3
      Top             =   2640
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   135
      Width           =   2865
   End
   Begin Crystal.CrystalReport oCrystalReport 
      Left            =   405
      Top             =   3645
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComCtl2.DTPicker DTHasta 
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      Top             =   2025
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   16842753
      CurrentDate     =   37518
   End
   Begin MSComCtl2.DTPicker DTDesde 
      Height          =   405
      Left            =   1920
      TabIndex        =   1
      Top             =   1515
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   714
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   16842753
      CurrentDate     =   37518
   End
   Begin ctrlayuda_f.Ctr_Ayuda Ctr_AYudaMoneda 
      Height          =   315
      Left            =   1845
      TabIndex        =   9
      Top             =   630
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
   Begin VB.Label lbl 
      Caption         =   "Tipo de Cambio :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   45
      TabIndex        =   11
      Top             =   1080
      Width           =   1875
   End
   Begin VB.Label lbl 
      Caption         =   "Moneda :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   675
      TabIndex        =   10
      Top             =   630
      Width           =   975
   End
   Begin VB.Label lbl 
      Caption         =   "Almacen"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   780
      TabIndex        =   7
      Top             =   180
      Width           =   975
   End
   Begin VB.Label lbl 
      Caption         =   "Desde"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   840
      TabIndex        =   6
      Top             =   1515
      Width           =   975
   End
   Begin VB.Label lbl 
      Caption         =   "Hasta"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   840
      TabIndex        =   5
      Top             =   1995
      Width           =   855
   End
End
Attribute VB_Name = "frmvtafacdet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adll As New dllgeneral.dll_general

Private Sub cmdAceptar_Click()
Dim Param(6) As Variant
Dim Form(2) As Variant


Param(0) = VGParamSistem.BDEmpresa
Param(1) = Left(Combo1.Text, 2)
Param(2) = DTDesde
Param(3) = DTHasta
Param(4) = TxtTc.Text
Param(5) = Ctr_AYudaMoneda.xclave

Form(0) = "@Empresa='" & VGParametros.nomempresa & "'"
Form(1) = "@ruc='" & VGParametros.RucEmpresa & "'"

Call ImpresionRptProc("vt_RepVtaFacDet.rpt", Form, Param, , "Ventas Factura Detallado")

End Sub

Private Sub Form_Load()
Call adll.llenacombo(Combo1, "select taalma,tadescri from tabalm where empresacodigo='" & VGParametros.empresacodigo & "' and puntovtacodigo='" & VGParametros.puntovta & "'", VGCNx)
Call Ctr_AYudaMoneda.Conexion(VGCNx)
DTDesde = Date
DTHasta = Date
TxtTc.Text = VGParamSistem.tipocambio

End Sub
