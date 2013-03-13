VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Begin VB.Form frmArtSinMov 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Artículos sin movimientos"
   ClientHeight    =   4275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4860
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   4860
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   2175
      Left            =   120
      TabIndex        =   8
      Top             =   1800
      Width           =   2895
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmArtSinMov.frx":0000
         Left            =   240
         List            =   "frmArtSinMov.frx":000A
         TabIndex        =   10
         Top             =   720
         Width           =   2055
      End
      Begin VB.ComboBox cmbOrden 
         Height          =   315
         ItemData        =   "frmArtSinMov.frx":0045
         Left            =   240
         List            =   "frmArtSinMov.frx":004F
         TabIndex        =   9
         Top             =   1560
         Width           =   2055
      End
      Begin VB.Label lblMoneda 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Filtrado por"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lblMoneda 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Ordenado por"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   11
         Top             =   1200
         Width           =   1095
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   2175
      Left            =   3120
      TabIndex        =   3
      Top             =   1800
      Width           =   1455
      Begin VB.CommandButton cmdOk 
         Caption         =   "Aceptar"
         Height          =   690
         Left            =   240
         Picture         =   "frmArtSinMov.frx":0068
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   945
      End
      Begin VB.CommandButton cmdQuit 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         Height          =   690
         Left            =   270
         Picture         =   "frmArtSinMov.frx":04AA
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1200
         Width           =   945
      End
   End
   Begin VB.Frame frmFecha 
      BackColor       =   &H00E0E0E0&
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4575
      Begin MSComCtl2.DTPicker dtpDesde 
         Height          =   315
         Left            =   1200
         TabIndex        =   1
         Top             =   840
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Format          =   62193665
         CurrentDate     =   37474
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuempresa 
         Height          =   315
         Left            =   1200
         TabIndex        =   7
         Top             =   360
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
         Requerido       =   0   'False
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   360
         TabIndex        =   6
         Top             =   480
         Width           =   735
      End
      Begin VB.Label lblDesde 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   360
         TabIndex        =   2
         Top             =   900
         Width           =   555
      End
   End
End
Attribute VB_Name = "frmArtSinMov"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdQuit_Click()
Unload Me
End Sub

Private Sub Form_Load()
AlinearFrm Me
dtpDesde.Value = Date - 180
dtpDesde = Fecha(1, dtpDesde)
Combo1.ListIndex = 1
cmbOrden.ListIndex = 0
Ctr_Ayuempresa.Conexion VGCNx
End Sub
Private Sub cmdOK_Click()
On Error GoTo Mensaje
Dim Form(7) As Variant
Dim aparam(5) As Variant
Form(0) = "desde = '" & dtpDesde & "'"
Form(1) = "filtro ='" & Combo1.text & "'"
Form(2) = "orden = " & cmbOrden.text
Form(3) = "empresa ='" & IIf(Ctr_Ayuempresa.xclave = "", "Todas", Ctr_Ayuempresa.xnombre) & "'"
aparam(0) = VGCNx.DefaultDatabase
aparam(1) = IIf(Ctr_Ayuempresa.xclave = "", "%%", Ctr_Ayuempresa.xclave)
aparam(2) = dtpDesde.Value
aparam(3) = Combo1.ListIndex
aparam(4) = cmbOrden.ListIndex
Call ImpresionRptProc("al_articulossinmovimientos.rpt", Form, aparam, , " Articulos sin movimientos ")
Exit Sub
Mensaje:
   Captura_error
   Screen.MousePointer = 1
   Exit Sub
End Sub

