VERSION 5.00
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Object = "{4D137D9C-00A6-4458-9B46-7E95DB76D55B}#9.0#0"; "TextFer.ocx"
Begin VB.Form frmRepClientes 
   Caption         =   "Reporte de Clientes"
   ClientHeight    =   3375
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5505
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3375
   ScaleWidth      =   5505
   StartUpPosition =   3  'Windows Default
   Begin TextFer.TxFer txtDistrito 
      Height          =   315
      Left            =   1395
      TabIndex        =   2
      Top             =   1545
      Width           =   3705
      _ExtentX        =   6535
      _ExtentY        =   556
      Object.CausesValidation=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   ""
      Valor           =   ""
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   2910
      TabIndex        =   4
      Top             =   2520
      Width           =   1320
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   1275
      TabIndex        =   3
      Top             =   2520
      Width           =   1320
   End
   Begin VB.ComboBox cbonegocio 
      Height          =   315
      Left            =   1395
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   1050
      Width           =   3705
   End
   Begin ctrlayuda_f.Ctr_Ayuda Ctr_Cliente 
      Height          =   300
      Left            =   1395
      TabIndex        =   0
      Top             =   630
      Width           =   3705
      _ExtentX        =   6535
      _ExtentY        =   529
      XcodMaxLongitud =   0
      xcodwith        =   600
      NomTabla        =   "vt_cliente"
      ListaCampos     =   "clientecodigo(1),clienterazonsocial(1)"
      XcodCampo       =   "clientecodigo"
      XListCampo      =   "clienterazonsocial"
      ListaCamposDescrip=   "Código,Razón_Social"
      ListaCamposText =   "clientecodigo,clienterazonsocial"
      Requerido       =   0   'False
   End
   Begin VB.Label Label2 
      Caption         =   "Distrito"
      Height          =   315
      Left            =   435
      TabIndex        =   7
      Top             =   1575
      Width           =   885
   End
   Begin VB.Label Label3 
      Caption         =   "Negocio"
      Height          =   315
      Left            =   435
      TabIndex        =   6
      Top             =   1110
      Width           =   885
   End
   Begin VB.Label Label1 
      Caption         =   "Codigo"
      Height          =   345
      Left            =   435
      TabIndex        =   5
      Top             =   675
      Width           =   885
   End
End
Attribute VB_Name = "frmRepClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
  Call Cargacombo
  Call Ctr_Cliente.conexion(VGCNx)
End Sub

Private Sub cmdAceptar_Click()
'FIXIT: Declare 'Aparam' con un tipo de datos de enlace en tiempo de compilación           FixIT90210ae-R1672-R1B8ZE
Dim Aparam(4) As Variant, Aformu(0) As Variant
Dim vgdll As New dllgeneral.dll_general

Aparam(0) = VGCNx.DefaultDatabase
Aparam(1) = IIf(RTrim$(Ctr_Cliente.xclave) = Empty, "%%", RTrim$(Ctr_Cliente.xclave))
Aparam(2) = IIf(cbonegocio.ListIndex < 0, "%%", cbonegocio.List(cbonegocio.ListIndex))
Aparam(3) = IIf(RTrim$(txtDistrito.Text) = Empty, "%%", RTrim$(txtDistrito.Text))

Call ImpresionRptProc(RutaRepProc & "RepccClientes.rpt", Aformu, Aparam)

End Sub

Private Sub cmdSalir_Click()
   Unload Me
End Sub

'FIXIT: Declare 'Cargacombo' con un tipo de datos de enlace en tiempo de compilación       FixIT90210ae-R1672-R1B8ZE
Function Cargacombo()
   Dim rscom As New ADODB.Recordset
   Dim J As Integer
   
   cbonegocio.Clear
   Set rscom = VGCNx.Execute("select * from vt_negocio")
   If rscom.RecordCount > 0 Then
     Do Until rscom.EOF
        cbonegocio.AddItem rscom!negociocodigo & "-" & rscom!negociodescripcion
        rscom.MoveNext
     Loop
   End If
   Set rscom = Nothing

End Function
