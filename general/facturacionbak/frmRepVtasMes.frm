VERSION 5.00
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmRepVtasMes 
   Caption         =   "Ventas Mensuales"
   ClientHeight    =   4650
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6150
   LinkTopic       =   "Form1"
   ScaleHeight     =   4650
   ScaleWidth      =   6150
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   3255
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   5415
      Begin MSComCtl2.DTPicker DTHasta 
         Height          =   375
         Left            =   1320
         TabIndex        =   3
         Top             =   2400
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
      Begin MSComCtl2.DTPicker DTDesde 
         Height          =   495
         Left            =   1320
         TabIndex        =   4
         Top             =   1680
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
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
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuempresa 
         Height          =   375
         Left            =   1320
         TabIndex        =   7
         Top             =   240
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   661
         XcodMaxLongitud =   3
         xcodwith        =   300
         NomTabla        =   "co_multiempresas"
         TituloAyuda     =   "Busqueda de Empresas"
         ListaCampos     =   "empresacodigo(1),empresadescripcion(1),agentederetencion(1)"
         XcodCampo       =   "empresacodigo"
         XListCampo      =   "empresadescripcion"
         ListaCamposDescrip=   "Codigo,Descripcion"
         ListaCamposText =   "empresacodigo,empresadescripcion,agentederetencion"
         Requerido       =   0   'False
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyudaMoneda 
         Height          =   375
         Left            =   1320
         TabIndex        =   9
         Top             =   720
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   661
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
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyuCliente 
         Height          =   375
         Left            =   1320
         TabIndex        =   11
         Top             =   1200
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   661
         XcodMaxLongitud =   11
         xcodwith        =   800
         NomTabla        =   "vt_Cliente"
         TituloAyuda     =   "Ayuda de Clientes"
         ListaCampos     =   $"frmRepVtasMes.frx":0000
         XcodCampo       =   "clientecodigo"
         XListCampo      =   "clienterazonsocial"
         ListaCamposDescrip=   "Codigo,Descripcion,Ruc,Direccion,Distrito,LimiteCred,Saldo,T,P,M,D"
         ListaCamposText =   $"frmRepVtasMes.frx":00E6
         Requerido       =   0   'False
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cliente "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   240
         TabIndex        =   12
         Top             =   1200
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Moneda :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Leempresa 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Empresa :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lbl 
         Caption         =   "Desde :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
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
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label lbl 
         Caption         =   "Hasta :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
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
         Top             =   2520
         Width           =   855
      End
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
      Left            =   1440
      TabIndex        =   1
      Top             =   3720
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
      Left            =   3000
      TabIndex        =   0
      Top             =   3720
      Width           =   1215
   End
   Begin Crystal.CrystalReport oCrystalReport 
      Left            =   120
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "frmRepVtasMes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adll As New dllgeneral.dll_general
''''''''''''''''''''''''''''''''''''''''''''''''''''
'Agregar:
Dim busca As New dll_apisgen.dll_apis
''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub cmdAceptar_Click()
Dim Param(7) As Variant
Dim formulas(3) As Variant
On Error GoTo Errores

If DTDesde > DTHasta Then
    MsgBox "Fecha Desde debe ser mayor a Fecha Hasta", vbInformation, "AVISO"
    Exit Sub
End If
If Ctr_AyudaMoneda.xclave = Empty Then
   MsgBox "ingrese moneda del reporte", vbInformation, "AVISO"
    Exit Sub
End If

Screen.MousePointer = 11
   
formulas(0) = "Emp='" & IIf(Ctr_Ayuempresa.xclave = Empty, "TODAS ", Ctr_Ayuempresa.xnombre) & "'"
formulas(1) = "moneda='" & RTrim(Ctr_AyudaMoneda.xnombre) & "'"
If Ctr_AyuCliente.xclave = "" Then
   formulas(2) = "cliente=' TODOS '"
 Else
   formulas(2) = "cliente='" & Ctr_AyuCliente.xnombre & "'"
End If

Param(0) = VGParamSistem.BDEmpresa
Param(1) = IIf(Ctr_Ayuempresa.xclave = Empty, "%%", Ctr_Ayuempresa.xclave)
Param(2) = Ctr_AyudaMoneda.xclave
Param(3) = DTDesde
Param(4) = DTHasta
Param(5) = VGcomputer
If Ctr_AyuCliente.xclave = "" Then
   Param(6) = "%%"
 Else
   Param(6) = Ctr_AyuCliente.xclave
End If


Call ImpresionRptProc("vt_rankingarticulosxmeses.rpt", formulas, Param, , "Ventas por Mensuales")
   

   
Screen.MousePointer = 1

Exit Sub
Errores:
  Screen.MousePointer = 1
  MsgBox "Error inesperado: " & Err.Number & "  " & Err.Description, vbExclamation
  Err = 0
  Exit Sub
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub




Private Sub DtDesde_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub DtHasta_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
    MostrarForm Me, "C2"
    DTDesde = Date
    DTHasta = Date
    Call Ctr_Ayuempresa.conexion(VGCNx)
    Call Ctr_AyudaMoneda.conexion(VGCNx)
    Call Ctr_AyuCliente.conexion(VGCNx)
End Sub

Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
     If KeyCode = 13 Then SendKeys "{TAB}"
End Sub


