VERSION 5.00
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form FrmRepRegVtas 
   Caption         =   "Registro de Ventas"
   ClientHeight    =   4500
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6690
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4500
   ScaleWidth      =   6690
   StartUpPosition =   3  'Windows Default
   Begin Crystal.CrystalReport oCrystalReport 
      Left            =   240
      Top             =   3000
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComCtl2.DTPicker DTHasta 
      Height          =   405
      Left            =   2445
      TabIndex        =   4
      Top             =   2940
      Width           =   1575
      _ExtentX        =   2778
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
      Format          =   97517569
      CurrentDate     =   37518
   End
   Begin MSComCtl2.DTPicker DTDesde 
      Height          =   405
      Left            =   2445
      TabIndex        =   3
      Top             =   2460
      Width           =   1575
      _ExtentX        =   2778
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
      Format          =   97517569
      CurrentDate     =   37518
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
      Left            =   2445
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   960
      Width           =   2610
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
      Left            =   1560
      TabIndex        =   5
      Top             =   3660
      Width           =   1215
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   4680
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   960
      Visible         =   0   'False
      Width           =   1305
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
      Index           =   1
      Left            =   3000
      TabIndex        =   6
      Top             =   3660
      Width           =   1215
   End
   Begin ctrlayuda_f.Ctr_Ayuda Ctr_Cliente 
      Height          =   375
      Left            =   2445
      TabIndex        =   1
      Top             =   1440
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   661
      XcodMaxLongitud =   11
      xcodwith        =   800
      NomTabla        =   "vt_cliente"
      TituloAyuda     =   "Ayuda de Clientes"
      ListaCampos     =   "clientecodigo(1),clienterazonsocial(1),clienteruc(1),clientedireccion(1),clientedistrito(1),clientesuspendido(1)"
      XcodCampo       =   "clientecodigo"
      XListCampo      =   "clienterazonsocial"
      ListaCamposDescrip=   "Codigo,RazonSocial,RUC,Direccion,Distrito,Suspendido"
      ListaCamposText =   "clientecodigo,clienterazonsocial,clienteruc,clientedireccion,clientedistrito,clientesuspendido"
      Requerido       =   0   'False
   End
   Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuempresa 
      Height          =   315
      Left            =   2475
      TabIndex        =   12
      Top             =   480
      Visible         =   0   'False
      Width           =   4050
      _ExtentX        =   7144
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
   End
   Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyudaMoneda 
      Height          =   315
      Left            =   2430
      TabIndex        =   2
      Top             =   1935
      Width           =   4050
      _ExtentX        =   7144
      _ExtentY        =   556
      XcodMaxLongitud =   3
      xcodwith        =   300
      NomTabla        =   "gr_moneda"
      TituloAyuda     =   "Busqueda de Moneda"
      ListaCampos     =   "monedacodigo(1),monedadescripcion(1)"
      XcodCampo       =   "monedacodigo"
      XListCampo      =   "monedadescripcion"
      ListaCamposDescrip=   "Codigo,Descripcion"
      ListaCamposText =   "monedacodigo,monedadescripcion"
      Requerido       =   0   'False
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Moneda :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1035
      TabIndex        =   14
      Top             =   1935
      Width           =   900
   End
   Begin VB.Label Lblempresa 
      AutoSize        =   -1  'True
      Caption         =   "Empresa :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1080
      TabIndex        =   13
      Top             =   480
      Visible         =   0   'False
      Width           =   960
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
      Left            =   1320
      TabIndex        =   11
      Top             =   2940
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
      Left            =   1320
      TabIndex        =   10
      Top             =   2460
      Width           =   855
   End
   Begin VB.Label lbl 
      Caption         =   "Punto de Venta"
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
      Index           =   0
      Left            =   480
      TabIndex        =   9
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label lbl 
      Caption         =   "Cliente"
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
      Left            =   1320
      TabIndex        =   8
      Top             =   1440
      Width           =   975
   End
End
Attribute VB_Name = "FrmRepRegVtas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adll As New dllgeneral.dll_general
Dim busca As New dll_apisgen.dll_apis

Private Sub cmdAceptar_Click(Index As Integer)
Dim arrform(6) As Variant
Dim arrparam(7) As Variant
 If DTDesde > DtHasta Then
     MsgBox "Fecha Desde debe ser mayor a Fecha Hasta", vbInformation, "AVISO"
     Exit Sub
 End If
       arrform(0) = "Empresa='" & VGParametros.nomempresa & "'"
        arrform(1) = "Desde='" & DTDesde & "'"
        arrform(2) = "Hasta='" & DtHasta & "'"
        If Combo1.ListIndex <> -1 Then
            arrform(3) = "PuntoVta='" & Combo1.Text & "'"
        Else
            arrform(3) = "PuntoVta='TODOS'"
        End If
        If Trim(Ctr_Cliente.xclave) <> "" Then
            arrform(4) = "Cliente='" & Ctr_Cliente.xnombre & "'"
        Else
            arrform(4) = "Cliente='TODOS'"
        End If
        
        If Ctr_AYudaMoneda.xclave <> "" Then
            arrform(5) = "Qmoneda='" & Ctr_AYudaMoneda.xnombre & "'"
        Else
            arrform(5) = "Qmoneda='TODOS'"
        End If
        
        arrparam(0) = VGCNx.DefaultDatabase
        arrparam(1) = VGParametros.empresacodigo
        arrparam(2) = IIf(Len(Trim(txt(0))) = 0, "%%", Trim(txt(0)))
        arrparam(3) = DTDesde
        arrparam(4) = DtHasta
        arrparam(5) = IIf(Trim(Ctr_Cliente.xclave) = "", "%%", Trim(Ctr_Cliente.xclave))
        arrparam(6) = IIf(Trim(Ctr_AYudaMoneda.xclave) = "", "%%", Trim(Ctr_AYudaMoneda.xclave))
        
       Call ImpresionRptProc("vt_RegVtas.rpt", arrform, arrparam, "", "Registro de Ventas")
 
End Sub

Private Sub cmdCancelar_Click(Index As Integer)
    Unload Me
End Sub

Private Sub Combo1_Click()
  If Combo1.ListCount > 0 Then
     txt(0) = adll.ComboDato(Combo1.Text)
  Else
     txt(0) = ""
  End If
End Sub

Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
Private Sub Ctr_Cliente_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub DtDesde_KeyDown(KeyCode As Integer, Shift As Integer)
      If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub DtHasta_KeyDown(KeyCode As Integer, Shift As Integer)
      If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
    MostrarFormVentas Me, "C2"
    Call adll.llenacombo(Combo1, "select puntovtacodigo,puntovtadescripcion from vt_puntoventa", VGCNx)
    Call Ctr_Cliente.conexion(VGCNx)
    Call Ctr_Ayuempresa.conexion(VGCNx)
    Call Ctr_AYudaMoneda.conexion(VGCNx)
    DTDesde = Fecha(1, VGParamSistem.FechaTrabajo)
    DtHasta = Fecha(2, VGParamSistem.FechaTrabajo)
End Sub

