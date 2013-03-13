VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Begin VB.Form FrmLibroBalancedeComprobacion 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Balance de Comprobacion"
   ClientHeight    =   5820
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7050
   LinkTopic       =   "Form1"
   ScaleHeight     =   5820
   ScaleWidth      =   7050
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Forma"
      Height          =   675
      Left            =   3465
      TabIndex        =   19
      Top             =   3495
      Width           =   3090
      Begin VB.OptionButton OptForma 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Mensual "
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   21
         Top             =   270
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.OptionButton OptForma 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Acumulado "
         Height          =   315
         Index           =   1
         Left            =   1665
         TabIndex        =   20
         Top             =   270
         Width           =   1260
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFC0&
      Height          =   1065
      Left            =   480
      TabIndex        =   16
      Top             =   360
      Width           =   6075
      Begin VB.CheckBox ChkCascada 
         BackColor       =   &H00FFFFC0&
         Caption         =   "En Cascada"
         Height          =   315
         Left            =   1290
         TabIndex        =   17
         Top             =   600
         Width           =   2040
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Balance de Comprobacion del mes Activo"
         Height          =   315
         Left            =   1290
         TabIndex        =   18
         Top             =   330
         Width           =   3645
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Tipo"
      Height          =   675
      Left            =   510
      TabIndex        =   13
      Top             =   3495
      Width           =   2910
      Begin VB.OptionButton OpTipo 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Historico"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   15
         Top             =   315
         Width           =   1245
      End
      Begin VB.OptionButton OpTipo 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Ajustado"
         Enabled         =   0   'False
         Height          =   195
         Index           =   1
         Left            =   1560
         TabIndex        =   14
         Top             =   315
         Width           =   990
      End
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   360
      Left            =   2175
      TabIndex        =   12
      Top             =   4755
      Width           =   1365
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   360
      Left            =   3525
      TabIndex        =   11
      Top             =   4755
      Width           =   1365
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFC0&
      Height          =   1080
      Left            =   480
      TabIndex        =   3
      Top             =   1425
      Width           =   6090
      Begin VB.ComboBox cmbNivel 
         Height          =   315
         Left            =   1035
         TabIndex        =   4
         Text            =   "Combo1"
         Top             =   225
         Width           =   1920
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   285
         Left            =   1050
         TabIndex        =   5
         Top             =   615
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "MM - MMMM"
         Format          =   64815107
         UpDown          =   -1  'True
         CurrentDate     =   37505
      End
      Begin ctrlayuda_f.Ctr_Ayuda CtrAyu_Moneda 
         Height          =   315
         Left            =   3825
         TabIndex        =   6
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
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         Caption         =   "Moneda :"
         Height          =   195
         Left            =   3120
         TabIndex        =   9
         Top             =   645
         Width           =   675
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         Caption         =   "Mes :"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   630
         Width           =   390
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         Caption         =   "Nivel :"
         Height          =   195
         Left            =   135
         TabIndex        =   7
         Top             =   270
         Width           =   450
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Filtro por Cuentas"
      Height          =   900
      Left            =   495
      TabIndex        =   0
      Top             =   2565
      Width           =   6075
      Begin VB.CheckBox ChkFiltcta 
         Height          =   300
         Left            =   1470
         TabIndex        =   1
         Top             =   -15
         Width           =   210
      End
      Begin ctrlayuda_f.Ctr_Ayuda CtrAyu_Cuenta 
         Height          =   315
         Left            =   120
         TabIndex        =   2
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
   Begin MSDataListLib.DataCombo DtCfiltro 
      Height          =   315
      Left            =   2235
      TabIndex        =   10
      Top             =   4290
      Width           =   4305
      _ExtentX        =   7594
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      ListField       =   "Descr"
      BoundColumn     =   "Cod"
      Text            =   "DataCombo1"
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "Filtro de Movimientos :"
      Height          =   195
      Left            =   555
      TabIndex        =   22
      Top             =   4335
      Width           =   1575
   End
End
Attribute VB_Name = "FrmLibroBalancedeComprobacion"
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
    Call CargaNivel
    Call CargaRsFiltro
    Call CtrAyu_Moneda.conexion(VGCNx)
    Call CtrAyu_Cuenta.conexion(VGCNx)
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
    Call ImpresionRptProc("ct_LibroBalancedeComprobacion.rpt", arrform, arrparm)
End Sub

Private Sub OptForma_Click(Index As Integer)
    lforma = Index + 1
End Sub

