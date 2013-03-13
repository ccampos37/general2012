VERSION 5.00
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Begin VB.Form frmRepImpresionLetras 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Impresión de Letras"
   ClientHeight    =   2115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5475
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   5475
   StartUpPosition =   3  'Windows Default
   Begin ctrlayuda_f.Ctr_Ayuda Ctr_NumDoc 
      Height          =   315
      Left            =   1260
      TabIndex        =   1
      Top             =   750
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   556
      XcodMaxLongitud =   0
      xcodwith        =   900
      NomTabla        =   "vt_cargo"
      ListaCampos     =   "documentocargo(1),cargonumdoc(1),clientecodigo(1)"
      XcodCampo       =   "cargonumdoc"
      XListCampo      =   "clientecodigo"
      ListaCamposDescrip=   "TD,NDoc,CodCli"
      ListaCamposText =   "documentocargo,cargonumdoc,clientecodigo"
      Requerido       =   0   'False
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   360
      Left            =   1425
      TabIndex        =   2
      Top             =   1395
      Width           =   1290
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   360
      Left            =   2760
      TabIndex        =   3
      Top             =   1395
      Width           =   1290
   End
   Begin ctrlayuda_f.Ctr_Ayuda Ctr_Doc 
      Height          =   300
      Left            =   1275
      TabIndex        =   0
      Top             =   405
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   529
      XcodMaxLongitud =   0
      xcodwith        =   500
      NomTabla        =   "cc_tipodocumento"
      ListaCampos     =   "tdocumentocodigo(1),tdocumentodescripcion(1)"
      XcodCampo       =   "tdocumentocodigo"
      XListCampo      =   "tdocumentodescripcion"
      ListaCamposDescrip=   "Código,Descripción"
      ListaCamposText =   "tdocumentocodigo,tdocumentodescripcion"
      Requerido       =   0   'False
   End
   Begin VB.Label Label2 
      Caption         =   "Nº Doc"
      Height          =   285
      Left            =   360
      TabIndex        =   5
      Top             =   795
      Width           =   765
   End
   Begin VB.Label Label1 
      Caption         =   "Documento"
      Height          =   240
      Left            =   345
      TabIndex        =   4
      Top             =   450
      Width           =   885
   End
End
Attribute VB_Name = "frmRepImpresionLetras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dllgeneral As New dllgeneral.dll_general

Private Sub cmdAceptar_Click()
  If Ctr_Doc.xclave = Empty Or Ctr_NumDoc.xclave = Empty Then
     MsgBox "Faltan seleccionar el número Documento y/o Tipo de Documento", vbInformation, Caption
     Exit Sub
  End If
  Call ImpresionLetra
End Sub

Private Sub Ctr_Doc_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
  Ctr_NumDoc.Filtro = "documentocargo='" & ColecCampos(0).Value & "' and isnull(cargoapeflgreg,0)<>1"
  Ctr_NumDoc.Ejecutar
End Sub

Private Sub Form_Load()
  Ctr_Doc.conexion cn
  Ctr_NumDoc.conexion cn
  Ctr_Doc.Filtro = "tdocumentotipo='C' and tdocumentodocrenovaletra='1'"
  Ctr_Doc.Ejecutar
End Sub

Private Sub cmdCancelar_Click()
  Unload Me
End Sub

Sub ImpresionLetra()
'FIXIT: Declare 'arrform' con un tipo de datos de enlace en tiempo de compilación          FixIT90210ae-R1672-R1B8ZE
Dim arrform(4) As Variant, arrparm(4) As Variant
Dim NombreRep As String
Dim mon As String
Dim cLetras As String
Dim rs As New ADODB.Recordset
Dim SQL As String
   Set rs = New ADODB.Recordset
   SQL = "select A.cargoapeimpape,A.monedacodigo from vt_cargo A where A.documentocargo='" & RTrim$(Ctr_Doc.xclave) & "' AND "
   SQL = SQL & "A.cargonumdoc='" & RTrim$(Ctr_NumDoc.xclave) & "' AND  A.clientecodigo like '" & RTrim$(Ctr_NumDoc.xnombre) & "'"
   Set rs = VGCNx.Execute(SQL)
   If rs.BOF Or rs.EOF Then
      MsgBox "No existen datos para el Documento", vbInformation, Caption
      Exit Sub
   End If
   
   cLetras = dllgeneral.NUMLET(Numero(Round(CDbl(rs(0)), 2))) & IIf(rs(1) = g_TipoSol, "Nuevos Soles", "Dolares Americanos")
   arrparm(0) = VGCNx.DefaultDatabase
   arrparm(1) = IIf(Ctr_Doc.xclave = Empty, "%", RTrim$(Ctr_Doc.xclave))
   arrparm(2) = IIf(Ctr_NumDoc.xclave = Empty, "%", RTrim$(Ctr_NumDoc.xclave))
   arrparm(3) = RTrim$(Ctr_NumDoc.xnombre)
   arrform(0) = "@Titulo='" & RTrim$(Ctr_NumDoc.xclave) & "'"
   arrform(1) = "@Empresa='" & g_DetalleEmpresa & "'"
   arrform(2) = "@LugarGiro='Los Olivos'"
   arrform(3) = "@Numletras= '" & cLetras & "'"
   
   NombreRep = "RepLetraimpresa.rpt"
   Call ImpresionRptProc(NombreRep, arrform, arrparm, Empty, "Planilla de " & "Impresión de Letras")
End Sub
