VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Begin VB.Form FrmVtasporPuntoVta 
   Caption         =   "Ventas por Punto de venta"
   ClientHeight    =   5055
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6975
   LinkTopic       =   "Form2"
   ScaleHeight     =   5055
   ScaleWidth      =   6975
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   1455
      Left            =   480
      TabIndex        =   8
      Top             =   3120
      Width           =   6135
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
         Height          =   855
         Index           =   0
         Left            =   960
         TabIndex        =   10
         Top             =   360
         Width           =   1575
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
         Height          =   855
         Index           =   1
         Left            =   3120
         TabIndex        =   9
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2655
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   6135
      Begin VB.CheckBox Check1 
         Caption         =   "Resumido"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   4680
         TabIndex        =   11
         Top             =   -120
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker DTHasta 
         Height          =   375
         Left            =   1680
         TabIndex        =   1
         Top             =   1920
         Width           =   1455
         _ExtentX        =   2566
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
         Format          =   98762753
         CurrentDate     =   37518
      End
      Begin MSComCtl2.DTPicker DTDesde 
         Height          =   375
         Left            =   1680
         TabIndex        =   2
         Top             =   1440
         Width           =   1455
         _ExtentX        =   2566
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
         Format          =   98762753
         CurrentDate     =   37518
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyuProducto 
         Height          =   375
         Left            =   1680
         TabIndex        =   3
         Top             =   960
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   661
         XcodMaxLongitud =   11
         xcodwith        =   800
         NomTabla        =   "maeart"
         TituloAyuda     =   "Busqueda de Codigos de producto"
         ListaCampos     =   "acodigo(1),adescri(1)"
         XcodCampo       =   "acodigo"
         XListCampo      =   "adescri"
         ListaCamposDescrip=   "Codigo,Descripcion"
         ListaCamposText =   "acodigo,adescri"
         Requerido       =   0   'False
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyuPuntoVta 
         Height          =   375
         Left            =   1680
         TabIndex        =   12
         Top             =   480
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   661
         XcodMaxLongitud =   2
         xcodwith        =   300
         NomTabla        =   "vt_puntoventa"
         TituloAyuda     =   "Busqueda de Punto de venta"
         ListaCampos     =   "puntovtacodigo(1),puntovtadescripcion(1)"
         XcodCampo       =   "puntovtacodigo"
         XListCampo      =   "puntovtadescripcion"
         ListaCamposDescrip=   "Codigo,Descripcion"
         ListaCamposText =   "puntovtacodigo,puntovtadescripcion"
         Requerido       =   0   'False
      End
      Begin VB.Label lbl 
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   7
         Top             =   1920
         Width           =   615
      End
      Begin VB.Label lbl 
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   6
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label lbl 
         Caption         =   "Punto de Venta"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label lbl 
         Caption         =   "Articulo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   4
         Top             =   960
         Width           =   855
      End
   End
End
Attribute VB_Name = "FrmVtasporPuntoVta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adll As New dllgeneral.dll_general
''''''''''''''''''''''''''''''''''''''''''''''''''''
'Agregar:
Dim busca As New dll_apisgen.dll_apis
Dim sresumen As Integer
''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub cmdAceptar_Click(Index As Integer)
Dim aform(3) As Variant
Dim aparam(7) As Variant

 If DTDesde > DTHasta Then
     MsgBox "Fecha Desde debe ser mayor a Fecha Hasta", vbInformation, "AVISO"
     Exit Sub
  End If
aform(0) = "Desde='" & DTDesde & "'"
aform(1) = "Hasta='" & DTHasta & "'"
If Ctr_AyuPuntoVta.xclave <> "" Then
   aform(2) = "puntoventa='" & Ctr_AyuPuntoVta.xclave & "'"
 Else
   aform(2) = "puntoventa='TODOS'"
End If
If Ctr_AyuProducto.xclave <> "" Then
   aform(3) = "Articulo='" & Ctr_AyuProducto.xclave & "'"
   aparam(3) = Ctr_AyuProducto.xclave
 Else
   aform(3) = "Articulo='TODOS'"
   aparam(3) = "%%"
End If
aparam(0) = VGCNx.DefaultDatabase
aparam(1) = VGParametros.empresacodigo
aparam(2) = IIf(Ctr_AyuPuntoVta.xclave = "", "%%", Ctr_AyuPuntoVta.xclave)
aparam(4) = DTDesde
aparam(5) = DTHasta
aparam(6) = 1
If Check1.Value = 1 Then
   Call ImpresionRptProc("vt_VtasxPtoVtaResumen.rpt", aform, aparam, , "Ventas por Punto de venta Resumen ")
Else
   Call ImpresionRptProc("vt_VtasxDocumentos.rpt", aform, aparam, , "Ventas por Punto de venta detallado")
End If
End Sub

Private Sub cmdCancelar_Click(Index As Integer)
    Unload Me
End Sub

Private Sub DtDesde_KeyDown(KeyCode As Integer, Shift As Integer)
      If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub DtHasta_KeyDown(KeyCode As Integer, Shift As Integer)
      If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
    Check1.Value = sresumen
    Call Ctr_AyuPuntoVta.Conexion(VGCNx)
    Call Ctr_AyuProducto.Conexion(VGCNx)
    DTDesde = Date
    DTHasta = Date
    If Check1.Value = 1 Then
       Ctr_AyuPuntoVta.Enabled = False
     Else
       Ctr_AyuPuntoVta.Enabled = True
    End If
    Ctr_AyuPuntoVta.filtro = " puntovtacodigo in (" & VGParametros.listaPuntoVtas & ")"
End Sub

Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub


Public Property Let resumen(ByVal pdato As Variant)
sresumen = pdato
End Property
