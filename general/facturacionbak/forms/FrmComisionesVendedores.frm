VERSION 5.00
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmComisionesVendedores 
   Caption         =   "Comisiones de Vendedores"
   ClientHeight    =   3420
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6510
   LinkTopic       =   "Form1"
   ScaleHeight     =   3420
   ScaleWidth      =   6510
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   2895
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6015
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
         Index           =   1
         Left            =   2760
         TabIndex        =   4
         Top             =   2040
         Width           =   1215
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
         Index           =   0
         Left            =   1320
         TabIndex        =   3
         Top             =   2040
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker DTPHasta 
         Height          =   375
         Left            =   1680
         TabIndex        =   1
         Top             =   1440
         Width           =   1575
         _ExtentX        =   2778
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
         Format          =   64815105
         CurrentDate     =   37518
      End
      Begin MSComCtl2.DTPicker DTPDesde 
         Height          =   375
         Left            =   1680
         TabIndex        =   2
         Top             =   840
         Width           =   1575
         _ExtentX        =   2778
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
         Format          =   64815105
         CurrentDate     =   37518
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyuVendedor 
         Height          =   375
         Left            =   1680
         TabIndex        =   8
         Top             =   360
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   661
         XcodMaxLongitud =   3
         xcodwith        =   200
         NomTabla        =   "vt_vendedor"
         TituloAyuda     =   "Ayuda de Vendedores"
         ListaCampos     =   "vendedorcodigo(1),vendedornombres(1)"
         XcodCampo       =   "vendedorcodigo"
         XListCampo      =   "vendedornombres"
         ListaCamposDescrip=   "Codigo,Descripcion"
         ListaCamposText =   "vendedorcodigo,vendedornombres"
         Requerido       =   0   'False
      End
      Begin VB.Label lbl 
         Caption         =   "Vendedor"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lbl 
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   6
         Top             =   960
         Width           =   855
      End
      Begin VB.Label lbl 
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   5
         Top             =   1560
         Width           =   735
      End
   End
End
Attribute VB_Name = "FrmComisionesVendedores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAceptar_Click(Index As Integer)
Dim aparam(5) As Variant
Dim aform(2) As Variant

aparam(0) = VGCNx.DefaultDatabase
aparam(1) = VGParametros.empresacodigo
aparam(2) = IIf(Ctr_AyuVendedor.xclave = "", "%%", Ctr_AyuVendedor.xclave)
aparam(3) = Format(DTPDesde.Value, "dd/mm/yyyy")
aparam(4) = Format(DTPHasta.Value, "dd/mm/yyyy")
aform(0) = "Desde='" & DTPDesde.Value & "'"
aform(1) = "Hasta='" & DTPHasta.Value & "'"
Call ImpresionRptProc("vt_comisionesVendedores.rpt", aform, aparam, , "Comisiones de Vendedores")
End Sub

Private Sub cmdCancelar_Click(Index As Integer)
Unload Me
End Sub

Private Sub Form_Load()
Call Ctr_AyuVendedor.Conexion(VGCNx)
DTPDesde = Fecha(1, VGParamSistem.FechaTrabajo)
DTPHasta = Fecha(2, VGParamSistem.FechaTrabajo)

End Sub
