VERSION 5.00
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Begin VB.Form FrmComprasxProveedor 
   Caption         =   "Informacion DAOT"
   ClientHeight    =   4080
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5790
   LinkTopic       =   "Form1"
   ScaleHeight     =   4080
   ScaleWidth      =   5790
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "IMPRIMIR POR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   240
      TabIndex        =   10
      Top             =   2640
      Width           =   2175
      Begin VB.CheckBox CheckRes 
         Caption         =   "Resumido"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   12
         Top             =   720
         Width           =   1215
      End
      Begin VB.CheckBox CheckDet 
         Caption         =   "Detallado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   11
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1335
      Left            =   2520
      TabIndex        =   5
      Top             =   2640
      Width           =   3015
      Begin VB.CommandButton cmdOK 
         Caption         =   "&Aceptar"
         Height          =   915
         Left            =   120
         Picture         =   "FrmComprasxProveedor.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   1260
      End
      Begin VB.CommandButton cmdBotones 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         Height          =   915
         Index           =   5
         Left            =   1560
         Picture         =   "FrmComprasxProveedor.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   1260
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2175
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   5295
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1200
         MaxLength       =   4
         TabIndex        =   4
         Top             =   960
         Width           =   615
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuempresa 
         Height          =   315
         Left            =   1200
         TabIndex        =   1
         Top             =   480
         Width           =   3855
         _ExtentX        =   6800
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
      Begin ctrlayuda_f.Ctr_Ayuda CtrAyu_Proveedor 
         Height          =   315
         Left            =   1155
         TabIndex        =   8
         Top             =   1440
         Width           =   3945
         _ExtentX        =   6959
         _ExtentY        =   556
         Enabled         =   0   'False
         XcodMaxLongitud =   11
         xcodwith        =   1000
         NomTabla        =   "cp_proveedor"
         TituloAyuda     =   "Busqueda de Proveedor"
         ListaCampos     =   "clientecodigo(1),clienterazonsocial(1),clienteruc(1),clientetelefono(1),proveedorcontribuyente(2)"
         XcodCampo       =   "clientecodigo"
         XListCampo      =   "clienterazonsocial"
         ListaCamposDescrip=   "Codigo,Descripcion"
         ListaCamposText =   "clientecodigo,clienterazonsocial,clienteruc,clientetelefono,proveedorcontribuyente"
      End
      Begin VB.Label Le_Proveedor 
         Caption         =   "Proveedor :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1485
         Width           =   1020
      End
      Begin VB.Label Label2 
         Caption         =   "Periodo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Empresa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   1215
      End
   End
End
Attribute VB_Name = "FrmComprasxProveedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBotones_Click(Index As Integer)
Unload Me
End Sub

Private Sub cmdOK_Click()
Dim aparam(5) As Variant
Dim aform(1) As Variant
If Ctr_Ayuempresa.xclave = "" Then
   MsgBox ("Ingrese Codigo de Empresa")
   Exit Sub
End If
If Text1.Text = "" Then
   MsgBox ("Ingrese Año a procesar")
   Exit Sub
End If
aform(0) = "periodo='" & Text1.Text & "'"

aparam(0) = VGCNx.DefaultDatabase
aparam(1) = Ctr_Ayuempresa.xclave
aparam(2) = Text1.Text
aparam(3) = IIf(CtrAyu_Proveedor.xclave = "", "%%", CtrAyu_Proveedor.xclave)
If CheckDet.Value = 1 Then
    aparam(4) = 0
    Call ImpresionRptProc("co_ComprasxProveedorDetallado.rpt", aform, aparam, , " Informacion DAO Detallado")
End If
If CheckRes.Value = 1 Then
    aparam(4) = 1
Call ImpresionRptProc("co_ComprasxProveedorresumen.rpt", aform, aparam, , " Informacion DAO")
End If
End Sub


Private Sub Form_Load()
Text1.Text = VGParamSistem.Anoproceso
Call Ctr_Ayuempresa.Conexion(VGCNx)
Call CtrAyu_Proveedor.Conexion(VGCNx)
End Sub
