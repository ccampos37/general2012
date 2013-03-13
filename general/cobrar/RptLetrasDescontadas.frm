VERSION 5.00
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "CONTROLAYUDA.OCX"
Begin VB.Form frmRepLetrasDescontadas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reporte Letras Descontadas"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5490
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   5490
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   390
      Index           =   0
      Left            =   1170
      TabIndex        =   2
      Top             =   2340
      Width           =   1380
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Cancelar"
      Height          =   390
      Index           =   1
      Left            =   2805
      TabIndex        =   3
      Top             =   2340
      Width           =   1380
   End
   Begin ctrlayuda_f.Ctr_Ayuda Ctr_Banco 
      Height          =   300
      Left            =   855
      TabIndex        =   1
      Top             =   1215
      Width           =   3600
      _ExtentX        =   6350
      _ExtentY        =   529
      XcodMaxLongitud =   0
      xcodwith        =   350
      NomTabla        =   "gr_banco"
      ListaCampos     =   "bancocodigo(1),bancodescripcion(1)"
      XcodCampo       =   "bancocodigo"
      XListCampo      =   "bancodescripcion"
      ListaCamposDescrip=   "Código,Descripción"
      ListaCamposText =   "bancocodigo,bancodescripcion"
      Requerido       =   0   'False
   End
   Begin ctrlayuda_f.Ctr_Ayuda Ctr_Cliente 
      Height          =   345
      Left            =   870
      TabIndex        =   0
      Top             =   750
      Width           =   4425
      _ExtentX        =   7805
      _ExtentY        =   609
      XcodMaxLongitud =   0
      xcodwith        =   900
      NomTabla        =   "vt_cliente"
      ListaCampos     =   "clientecodigo(1),clienterazonsocial(1)"
      XcodCampo       =   "clientecodigo"
      XListCampo      =   "clienterazonsocial"
      ListaCamposDescrip=   "Código,Descripción"
      ListaCamposText =   "clientecodigo,clienterazonsocial"
      Requerido       =   0   'False
   End
   Begin VB.Label Label2 
      Caption         =   "Bancos"
      Height          =   225
      Left            =   150
      TabIndex        =   5
      Top             =   1275
      Width           =   690
   End
   Begin VB.Label Label1 
      Caption         =   "Clientes"
      Height          =   255
      Left            =   180
      TabIndex        =   4
      Top             =   795
      Width           =   660
   End
End
Attribute VB_Name = "frmRepLetrasDescontadas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
  Call Ctr_Banco.conexion(cn)
  Call Ctr_Cliente.conexion(cn)
End Sub

Private Sub cmdAceptar_Click(Index As Integer)
  Select Case Index
    Case 0:
      Call Imprimir
    
    Case 1:
      Unload Me
  End Select

End Sub

Sub Imprimir()
Dim arrform(0) As Variant, arrparm(3) As Variant
Dim NombreRep As String, CadOrden As String
Dim NombrePC As String
Dim ValorRango As String
Dim i As Integer
   Randomize   ' Inicializa el generador de números aleatorios.
   NombrePC = Trim(Str(CLng(Rnd * 10000000)))
   arrparm(0) = cn.DefaultDatabase
   arrparm(1) = IIf(Ctr_Banco.xclave = Empty, "%", Trim(Ctr_Banco.xclave))
   arrparm(2) = IIf(Ctr_Cliente.xclave = Empty, "%", Trim(Ctr_Cliente.xclave))
   NombreRep = RutaRepProc & "RepccLetrasBancos.rpt"
   CadOrden = ""
   Call ImpresionRptProc(NombreRep, arrform, arrparm, , "Reporte de Letras Descontadas")
End Sub
