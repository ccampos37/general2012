VERSION 5.00
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Begin VB.Form frmCtasxrendir 
   Caption         =   "Cuentas por Rendir"
   ClientHeight    =   5355
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6555
   LinkTopic       =   "Form1"
   ScaleHeight     =   5355
   ScaleWidth      =   6555
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   1695
      Left            =   360
      TabIndex        =   10
      Top             =   3000
      Width           =   5655
      Begin VB.CommandButton cmdimprimir 
         Caption         =   "&Imprimir"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Index           =   3
         Left            =   960
         Picture         =   "frmCtasxrendir.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton cmdSalir 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         Height          =   915
         Index           =   5
         Left            =   2865
         Picture         =   "frmCtasxrendir.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   480
         Width           =   1260
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2655
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   5655
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyudaCaja 
         Height          =   375
         Left            =   1320
         TabIndex        =   2
         Top             =   1080
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   661
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
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuempresa 
         Height          =   375
         Left            =   1320
         TabIndex        =   1
         Top             =   480
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   661
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
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyuEntidad 
         Height          =   375
         Left            =   1320
         TabIndex        =   3
         Top             =   1560
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   661
         XcodMaxLongitud =   0
         xcodwith        =   1000
         NomTabla        =   "ct_entidad"
         ListaCampos     =   "entidadcodigo(1),entidadrazonsocial(1)"
         XcodCampo       =   "entidadcodigo"
         XListCampo      =   "entidadrazonsocial"
         ListaCamposDescrip=   "Código,Descripción"
         ListaCamposText =   "entidadcodigo,entidadrazonsocial"
         Requerido       =   0   'False
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayutransf 
         Height          =   315
         Left            =   1320
         TabIndex        =   4
         Top             =   2040
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   556
         XcodMaxLongitud =   7
         xcodwith        =   800
         NomTabla        =   "te_cabecerarecibos"
         TituloAyuda     =   "Busqueda de Documentos x rendir"
         ListaCampos     =   "cabrec_numreciboegreso(1),cabrec_descripcion(1),SaldoDocxRendir(1),clientecodigo(1)"
         XcodCampo       =   "cabrec_numreciboegreso"
         XListCampo      =   "cabrec_descripcion"
         ListaCamposDescrip=   "Nro.transferencia,descripcion,Saldo,usuario"
         ListaCamposText =   "cabrec_numreciboegreso,cabrec_descripcion,SaldoDocxRendir,clientecodigo"
      End
      Begin VB.Label LeReferencia 
         AutoSize        =   -1  'True
         Caption         =   "Nro.Transf."
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
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Responsable"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   120
         TabIndex        =   8
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Lblempresa 
         AutoSize        =   -1  'True
         Caption         =   "Empresa :"
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
         TabIndex        =   7
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Cod. Caja"
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
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   1080
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmCtasxrendir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdimprimir_Click(Index As Integer)
Dim aparam(3) As Variant
Dim aform(1) As Variant
aparam(0) = VGCNx.DefaultDatabase
aparam(1) = Ctr_Ayuempresa.xclave
aparam(2) = Ctr_Ayutransf.xclave
aform(0) = "transfer='DOCUMENTO : " & Ctr_Ayutransf.xclave & "'"
Call ImpresionRptProc("te_ctasxrendir.rpt", aform, aparam, , "Cuentas por rendir ")

End Sub

Private Sub cmdSalir_Click(Index As Integer)
Unload Me
End Sub

Private Sub Ctr_AyuEntidad_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
Ctr_Ayutransf.Filtro = "empresacodigo='" & Ctr_Ayuempresa.xclave & "' and cajacodigo='" & Ctr_AyudaCaja.xclave & "' and clientecodigo='" & Ctr_AyuEntidad.xclave & "'"
End Sub

Private Sub Form_Load()
Call Ctr_Ayuempresa.conexion(VGCNx)
Call Ctr_AyudaCaja.conexion(VGCNx)
Call Ctr_AyuEntidad.conexion(VGCNx)
Call Ctr_Ayutransf.conexion(VGCNx)
If VGParametros.sistemamultiempresas = False Then
   Ctr_Ayuempresa.xclave = VGParametros.empresacodigo: Ctr_Ayuempresa.Ejecutar
   Ctr_Ayuempresa.Enabled = False
End If

End Sub

