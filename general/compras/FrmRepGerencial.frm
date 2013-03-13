VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Begin VB.Form FrmRepGerencial 
   Caption         =   "Informes Gerenciales"
   ClientHeight    =   4290
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7725
   LinkTopic       =   "Form1"
   ScaleHeight     =   4290
   ScaleWidth      =   7725
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Listar  Por"
      Height          =   735
      Left            =   0
      TabIndex        =   17
      Top             =   2520
      Width           =   7635
      Begin VB.OptionButton OptCCostos 
         Caption         =   "Centro de Costos"
         Height          =   492
         Left            =   1560
         TabIndex        =   4
         Top             =   120
         Width           =   1815
      End
      Begin VB.OptionButton Optdetalle 
         Caption         =   "Mensualizado"
         Height          =   492
         Left            =   4440
         TabIndex        =   5
         Top             =   120
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2355
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   7635
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyuRendicion 
         Height          =   315
         Left            =   1500
         TabIndex        =   1
         Top             =   855
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   556
         XcodMaxLongitud =   20
         xcodwith        =   1000
         NomTabla        =   "te_rendiciones"
         TituloAyuda     =   "Busqueda de Nro de rendicion"
         ListaCampos     =   "rendicionnumero(1),monedacodigo(1),rendicionfecha(2)"
         XcodCampo       =   "rendicionnumero"
         XListCampo      =   "monedacodigo"
         ListaCamposDescrip=   "NroRendicion,Moneda"
         ListaCamposText =   "rendicionnumero,monedacodigo"
         Requerido       =   0   'False
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyuEntidad 
         Height          =   315
         Left            =   1500
         TabIndex        =   3
         Top             =   1800
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   556
         XcodMaxLongitud =   11
         xcodwith        =   1000
         NomTabla        =   "ct_entidad"
         TituloAyuda     =   "Busqueda de Entidad"
         ListaCampos     =   "entidadcodigo(1),entidadrazonsocial(1)"
         XcodCampo       =   "entidadcodigo"
         XListCampo      =   "entidadrazonsocial"
         ListaCamposDescrip=   "Código,Descripción"
         ListaCamposText =   "entidadcodigo,entidadrazonsocial"
         Requerido       =   0   'False
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyuOficina 
         Height          =   300
         Left            =   1500
         TabIndex        =   0
         Top             =   360
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   529
         XcodMaxLongitud =   0
         xcodwith        =   1000
         NomTabla        =   "cp_oficina"
         TituloAyuda     =   "Ayuda de Tipo Analitico"
         ListaCampos     =   "vendedorcodigo(1),vendedornombres(1)"
         XcodCampo       =   "vendedorcodigo"
         XListCampo      =   "vendedornombres"
         ListaCamposDescrip=   "Código,Descripción"
         ListaCamposText =   "vendedorcodigo,vendedornombres"
         Requerido       =   0   'False
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyuCcosto 
         Height          =   315
         Left            =   1500
         TabIndex        =   2
         Top             =   1320
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   556
         XcodMaxLongitud =   10
         xcodwith        =   900
         NomTabla        =   "ct_centrocosto"
         TituloAyuda     =   "Busqueda de Centro de Costos"
         ListaCampos     =   "centrocostocodigo(1),centrocostodescripcion(1)"
         XcodCampo       =   "centrocostocodigo"
         XListCampo      =   "centrocostodescripcion"
         ListaCamposDescrip=   "Código,Descripción"
         ListaCamposText =   "centrocostocodigo,centrocostodescripcion"
         Requerido       =   0   'False
      End
      Begin VB.Label lbccosto 
         AutoSize        =   -1  'True
         Caption         =   "C.Costo"
         Height          =   195
         Left            =   360
         TabIndex        =   19
         Top             =   1365
         Width           =   795
      End
      Begin VB.Label Leoficina 
         Caption         =   "Oficina :"
         Height          =   255
         Left            =   360
         TabIndex        =   18
         Top             =   375
         Width           =   840
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Rendicion :"
         Height          =   195
         Left            =   345
         TabIndex        =   16
         Top             =   915
         Width           =   810
      End
      Begin VB.Label Lblanalitico 
         AutoSize        =   -1  'True
         Caption         =   "Entidad"
         Height          =   195
         Left            =   360
         TabIndex        =   15
         Top             =   1845
         Width           =   540
      End
   End
   Begin VB.Frame Frame2 
      Height          =   765
      Left            =   0
      TabIndex        =   10
      Top             =   3360
      Width           =   6240
      Begin VB.CheckBox ChkFech 
         Caption         =   "Rango de Fechas"
         Height          =   285
         Left            =   75
         TabIndex        =   6
         Top             =   -45
         Width           =   1620
      End
      Begin MSComCtl2.DTPicker DTPFechaIni 
         Height          =   300
         Left            =   1260
         TabIndex        =   7
         Top             =   315
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   53411841
         CurrentDate     =   37623.1285069444
      End
      Begin MSComCtl2.DTPicker DTPFechaFin 
         Height          =   300
         Left            =   4140
         TabIndex        =   8
         Top             =   315
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   53411841
         CurrentDate     =   37623.1264351852
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Inicio :"
         Height          =   210
         Left            =   150
         TabIndex        =   13
         Top             =   360
         Width           =   1110
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha Fin :"
         Height          =   210
         Left            =   3195
         TabIndex        =   12
         Top             =   360
         Width           =   810
      End
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   330
      Left            =   6360
      TabIndex        =   9
      Top             =   3420
      Width           =   1260
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   330
      Left            =   6375
      TabIndex        =   11
      Top             =   3810
      Width           =   1260
   End
End
Attribute VB_Name = "FrmRepGerencial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ChkFech_Click()
If ChkFech.Value = 1 Then
    DTPFechaIni.Enabled = True
    DTPFechaFin.Enabled = True
  Else
    DTPFechaIni.Enabled = False
    DTPFechaFin.Enabled = False
End If
End Sub
Private Sub chkflagmodo_Click()
    If chkflagmodo.Value = 1 Then
        ChkCtaCte.Enabled = True: ChkActCaja.Enabled = True: ChkRegComp.Enabled = True
      Else
        ChkCtaCte.Enabled = False: ChkActCaja.Enabled = False: ChkRegComp.Enabled = False
    End If
End Sub
Private Sub CmdAceptar_Click()
    Call imprimir
End Sub
Private Sub CmdCancelar_Click()
    Unload Me
End Sub
Private Sub Ctr_AyuProveedor_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
   Ctr_AyuEntidad.Enabled = False
End Sub
Private Sub Form_Load()
    Call Ctr_AyuOficina.Conexion(VGcnx)
    Call Ctr_AyuRendicion.Conexion(VGcnx)
    Call Ctr_AyuCcosto.Conexion(VGcnx)
    Call Ctr_AyuEntidad.Conexion(VGcnxCT)
 '  Optresumen.Value = 1
    DTPFechaIni = Date
    DTPFechaFin = Date
End Sub
Public Sub imprimir()
Dim arrform(0) As Variant, arrparm(9) As Variant
On Error GoTo Imprime
    Screen.MousePointer = 11
    '@BaseCompra, @BaseConta, @Prove, @Ano, @flagfecha, @Fechaini, @fechafin, @cuenta
    Dim rsdate As New ADODB.Recordset
    Set VGvardllgen = New dllgeneral.dll_general
    arrparm(0) = VGParamSistem.BDEmpresa
    arrparm(1) = VGParamSistem.BDEmpresa
    arrparm(2) = VGParamSistem.Anoproceso
    arrparm(3) = FechS(DTPFechaIni.Value, 1)
    arrparm(4) = FechS(DTPFechaFin.Value, 1)
    arrparm(5) = 1
    arrparm(6) = IIf(Trim(Ctr_Ayugastos.xclave) = "", "%%", Trim(Ctr_Ayugastos.xclave))
    arrparm(7) = "%%"
    arrparm(8) = "%%"
    If Trim(Ctr_AyuProveedor.xclave) <> "" Then
       arrparm(5) = 1
       arrparm(7) = Ctr_AyuProveedor.xclave
       arrparm(8) = "%%"
    End If
    If Trim(Ctr_AyuEntidad.xclave) <> "" Then
       arrparm(5) = 2
       arrparm(7) = "%%"
       arrparm(8) = Ctr_AyuEntidad.xclave
    End If
    
    If Optresumen.Value Then
       Call ImpresionRptProc("co_gastosResumen.rpt", arrform, arrparm, , "Gastos Resumidos ")
     Else
       Call ImpresionRptProc("co_gastosDetallado.rpt", arrform, arrparm, , "Gastos Detallados ")
    End If
    Screen.MousePointer = 1
    Exit Sub
Imprime:
Screen.MousePointer = 1
MsgBox Err.Description
End Sub



