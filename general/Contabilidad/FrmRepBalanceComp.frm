VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Begin VB.Form FrmRepBalanceComp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reporte Balance de Comprobancion"
   ClientHeight    =   5505
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6600
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   6600
   Begin VB.Frame Frame5 
      Caption         =   "Filtro por Cuentas"
      Height          =   900
      Left            =   75
      TabIndex        =   20
      Top             =   2220
      Width           =   6075
      Begin VB.CheckBox ChkFiltcta 
         Height          =   300
         Left            =   1470
         TabIndex        =   21
         Top             =   -15
         Width           =   210
      End
      Begin ctrlayuda_f.Ctr_Ayuda CtrAyu_Cuenta 
         Height          =   315
         Left            =   120
         TabIndex        =   22
         Top             =   375
         Width           =   5850
         _ExtentX        =   10319
         _ExtentY        =   556
         Enabled         =   0   'False
         XcodMaxLongitud =   20
         xcodwith        =   1000
         NomTabla        =   "ct_cuenta"
         TituloAyuda     =   "Busqueda de Cuenta"
         ListaCampos     =   "cuentacodigo(1),cuentadescripcion(1),cuentaestadoccostos(2),cuentaestadoanalitico(2),cuentadocumento(2),tipoanaliticocodigo(1)"
         XcodCampo       =   "cuentacodigo"
         XListCampo      =   "cuentadescripcion"
         ListaCamposDescrip=   "Código,Descripción"
         ListaCamposText =   "cuentacodigo,cuentadescripcion,cuentaestadoccostos,cuentaestadoanalitico,cuentadocumento,tipoanaliticocodigo"
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1080
      Left            =   60
      TabIndex        =   13
      Top             =   1080
      Width           =   6090
      Begin VB.ComboBox cmbNivel 
         Height          =   315
         Left            =   1035
         TabIndex        =   15
         Text            =   "Combo1"
         Top             =   225
         Width           =   1920
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   285
         Left            =   1050
         TabIndex        =   14
         Top             =   615
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "MM - MMMM"
         Format          =   4063235
         UpDown          =   -1  'True
         CurrentDate     =   37505
      End
      Begin ctrlayuda_f.Ctr_Ayuda CtrAyu_Moneda 
         Height          =   315
         Left            =   3825
         TabIndex        =   16
         Top             =   615
         Width           =   2160
         _ExtentX        =   3810
         _ExtentY        =   556
         XcodMaxLongitud =   2
         NomTabla        =   "gr_moneda"
         TituloAyuda     =   "Busqueda de Moneda"
         ListaCampos     =   "monedacodigo(1),monedadescripcion(1)"
         XcodCampo       =   "monedacodigo"
         XListCampo      =   "monedadescripcion"
         ListaCamposDescrip=   "Código,Descripción"
         ListaCamposText =   "monedacodigo,monedadescripcion"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nivel :"
         Height          =   195
         Left            =   135
         TabIndex        =   19
         Top             =   270
         Width           =   450
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Mes :"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   630
         Width           =   390
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Moneda :"
         Height          =   195
         Left            =   3120
         TabIndex        =   17
         Top             =   645
         Width           =   675
      End
   End
   Begin MSDataListLib.DataCombo DtCfiltro 
      Height          =   315
      Left            =   1815
      TabIndex        =   11
      Top             =   3945
      Width           =   4305
      _ExtentX        =   7594
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      ListField       =   "Descr"
      BoundColumn     =   "Cod"
      Text            =   "DataCombo1"
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   360
      Left            =   3101
      TabIndex        =   10
      Top             =   4410
      Width           =   1365
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   360
      Left            =   1758
      TabIndex        =   9
      Top             =   4410
      Width           =   1365
   End
   Begin VB.Frame Frame3 
      Caption         =   "Tipo"
      Height          =   675
      Left            =   90
      TabIndex        =   6
      Top             =   3150
      Width           =   2910
      Begin VB.OptionButton OpTipo 
         Caption         =   "Ajustado"
         Enabled         =   0   'False
         Height          =   195
         Index           =   1
         Left            =   1560
         TabIndex        =   8
         Top             =   315
         Width           =   990
      End
      Begin VB.OptionButton OpTipo 
         Caption         =   "Historico"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   315
         Width           =   1245
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1065
      Left            =   60
      TabIndex        =   3
      Top             =   15
      Width           =   6075
      Begin VB.CheckBox ChkCascada 
         Caption         =   "En Cascada"
         Height          =   315
         Left            =   1290
         TabIndex        =   5
         Top             =   600
         Width           =   2040
      End
      Begin VB.Label Label2 
         Caption         =   "Balance de Comprobacion del mes Activo"
         Height          =   315
         Left            =   1290
         TabIndex        =   4
         Top             =   330
         Width           =   3645
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Forma"
      Height          =   675
      Left            =   3045
      TabIndex        =   0
      Top             =   3150
      Width           =   3090
      Begin VB.OptionButton OptForma 
         Caption         =   "Acumulado "
         Height          =   315
         Index           =   1
         Left            =   1665
         TabIndex        =   2
         Top             =   270
         Width           =   1260
      End
      Begin VB.OptionButton OptForma 
         Caption         =   "Mensual "
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   270
         Visible         =   0   'False
         Width           =   945
      End
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Filtro de Movimientos :"
      Height          =   195
      Left            =   135
      TabIndex        =   12
      Top             =   3990
      Width           =   1575
   End
End
Attribute VB_Name = "FrmRepBalanceComp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Dim RsFiltro As ADODB.Recordset
Dim lforma As Integer

Private Sub ChkFiltcta_Click()
    If ChkFiltcta.Value = 1 Then
        CtrAyu_Cuenta.Enabled = True
      Else
        CtrAyu_Cuenta.Enabled = False
     End If
End Sub
Private Sub cmdAceptar_Click()
    Screen.MousePointer = 11
    Call imprimir
    Screen.MousePointer = 1
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Height = 5235
    Width = 6315
    Call CargaNivel
    Call CargaRsFiltro
    Call CtrAyu_Moneda.Conexion(VGCNx)
    Call CtrAyu_Cuenta.Conexion(VGCNx)
    CtrAyu_Moneda.xclave = VGParametros.monedabase: CtrAyu_Moneda.Ejecutar
    OptForma(1).Value = True:
    OpTipo(0).Value = True
    DTPicker1.Value = DateSerial(CInt(VGParamSistem.Anoproceso), CInt(VGParamSistem.Mesproceso), 1)
End Sub
Private Sub CargaNivel()
    Dim i As Integer
    For i = 1 To VGnumnivelescuenta
        cmbNivel.AddItem Format(i, "0")
    Next
    cmbNivel.ListIndex = 0
End Sub
Private Sub CargaRsFiltro()
    Set RsFiltro = New ADODB.Recordset
    RsFiltro.Fields.Append "Cod", adVarChar, 2
    RsFiltro.Fields.Append "Descr", adVarChar, 50
    RsFiltro.Open
    RsFiltro.AddNew
    RsFiltro!Cod = "0"
    RsFiltro!Descr = "Todos las cuentas"
    RsFiltro.Update
    RsFiltro.AddNew
    RsFiltro!Cod = "1"
    RsFiltro!Descr = "Cuenta con Movimientos y Saldos Acumulados"
    RsFiltro.Update
    Set DtCfiltro.RowSource = RsFiltro
    DtCfiltro.BoundText = "1"
End Sub
Public Sub imprimir()
'FIXIT: Declare 'arrform' con un tipo de datos de enlace en tiempo de compilación          FixIT90210ae-R1672-R1B8ZE
Dim arrform(2) As Variant, arrparm(11) As Variant
Dim mon As String
    If CtrAyu_Moneda.xclave = "01" Then
        mon = 1
     Else
        mon = 2
    End If
     '@Base, @Anno, @Mes, @Nivel, @NoEnCascada, @Corden, @opvista
    arrparm(0) = VGParamSistem.BDEmpresa
    arrparm(1) = VGParametros.empresacodigo
    arrparm(2) = VGParamSistem.Anoproceso
    arrparm(3) = Format(Month(DTPicker1), "0") 'VGParamSistem.Mesproceso
    arrparm(4) = CInt(cmbNivel.Text)
    arrparm(5) = IIf(ChkCascada.Value = 1, 0, 1)
    arrparm(6) = "left(Cuenta,2) Asc,Nivel Desc"
    arrparm(7) = CInt(DtCfiltro.BoundText)
    If ChkFiltcta.Value = 1 Then
        If CtrAyu_Cuenta.xclave = "" Then
            MsgBox "Debe escoger una cuenta ", vbInformation
            Exit Sub
        End If
        arrparm(8) = Format(Len(Trim$(CtrAyu_Cuenta.xclave)), "0")
        arrparm(9) = Trim$(CtrAyu_Cuenta.xclave) & "%"
      Else
        arrparm(8) = 0
        arrparm(9) = 0
    End If
    arrparm(10) = IIf(VGParametros.sistemamonista, "1", "0")
    arrform(0) = "pmon=" & mon
    arrform(1) = "pforma=" & lforma
    'Call ImpresionRptProc("rptBalanceComprob.rpt", arrform, arrparm)
    Call ImpresionRptProc("ct_BalanceComprobacion.rpt", arrform, arrparm)
End Sub

Private Sub OptForma_Click(Index As Integer)
    lforma = Index + 1
End Sub
