VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Begin VB.Form Frmreplistgastos 
   Caption         =   "Form1"
   ClientHeight    =   5850
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7050
   LinkTopic       =   "Form1"
   ScaleHeight     =   5850
   ScaleWidth      =   7050
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame5 
      Height          =   1335
      Left            =   3600
      TabIndex        =   30
      Top             =   4365
      Width           =   1695
      Begin VB.CheckBox Chkdetraccion 
         Caption         =   "Solo Detraccion"
         Height          =   495
         Left            =   240
         TabIndex        =   31
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Lista Por"
      Height          =   1455
      Left            =   5040
      TabIndex        =   26
      Top             =   2805
      Width           =   1695
      Begin VB.OptionButton Option1 
         Caption         =   "Analitico"
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton Optionproveedor 
         Caption         =   "Proveedor"
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   450
      Left            =   5400
      TabIndex        =   16
      Top             =   5205
      Width           =   1260
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   450
      Left            =   5400
      TabIndex        =   15
      Top             =   4605
      Width           =   1260
   End
   Begin VB.Frame Frame2 
      Height          =   1365
      Left            =   240
      TabIndex        =   9
      Top             =   4365
      Width           =   3240
      Begin VB.CheckBox ChkFech 
         Caption         =   "Rango de Fechas"
         Height          =   285
         Left            =   75
         TabIndex        =   10
         Top             =   -45
         Width           =   1620
      End
      Begin MSComCtl2.DTPicker DTPFechaIniX 
         Height          =   300
         Left            =   -1920
         TabIndex        =   11
         Top             =   315
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   50724865
         CurrentDate     =   37623.1285069444
      End
      Begin MSComCtl2.DTPicker DTPFechaFin 
         Height          =   300
         Left            =   1260
         TabIndex        =   12
         Top             =   675
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   50724865
         CurrentDate     =   37623.1264351852
      End
      Begin MSComCtl2.DTPicker DTPFechaIni 
         Height          =   300
         Left            =   1320
         TabIndex        =   29
         Top             =   240
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   50724865
         CurrentDate     =   37623.1264351852
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha Fin :"
         Height          =   210
         Left            =   315
         TabIndex        =   14
         Top             =   720
         Width           =   810
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Inicio :"
         Height          =   210
         Left            =   150
         TabIndex        =   13
         Top             =   360
         Width           =   1110
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2565
      Left            =   240
      TabIndex        =   1
      Top             =   135
      Width           =   6555
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayugastos 
         Height          =   315
         Left            =   1545
         TabIndex        =   3
         Top             =   615
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   556
         XcodMaxLongitud =   20
         xcodwith        =   1000
         NomTabla        =   "co_gastos"
         TituloAyuda     =   "Busqueda de Cuenta de Gastos"
         ListaCampos     =   "gastoscodigo(1),gastosdescripcion(1)"
         XcodCampo       =   "gastoscodigo"
         XListCampo      =   "gastosdescripcion"
         ListaCamposDescrip=   "Código,Descripción"
         ListaCamposText =   "gastoscodigo,gastosdescripcion"
         Requerido       =   0   'False
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyuProveedor 
         Height          =   315
         Left            =   1545
         TabIndex        =   4
         Top             =   960
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   556
         XcodMaxLongitud =   11
         xcodwith        =   1000
         NomTabla        =   "cp_proveedor"
         TituloAyuda     =   "Busqueda de Proveedor"
         ListaCampos     =   "clientecodigo(1),clienterazonsocial(1),clienteruc(1),clientetelefono(1)"
         XcodCampo       =   "clientecodigo"
         XListCampo      =   "clienterazonsocial"
         ListaCamposDescrip=   "Codigo,Descripcion"
         ListaCamposText =   "clientecodigo,clienterazonsocial,clienteruc,clientetelefono"
         Requerido       =   0   'False
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyuEntidad 
         Height          =   315
         Left            =   1545
         TabIndex        =   5
         Top             =   1320
         Width           =   4815
         _ExtentX        =   8493
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
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyuCcosto 
         Height          =   315
         Left            =   1545
         TabIndex        =   6
         Top             =   1680
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   556
         XcodMaxLongitud =   10
         xcodwith        =   1000
         NomTabla        =   "ct_centrocosto"
         TituloAyuda     =   "Busqueda de Centro de Costos"
         ListaCampos     =   "centrocostocodigo(1),centrocostodescripcion(1)"
         XcodCampo       =   "centrocostocodigo"
         XListCampo      =   "centrocostodescripcion"
         ListaCamposDescrip=   "Código,Descripción"
         ListaCamposText =   "centrocostocodigo,centrocostodescripcion"
         Requerido       =   0   'False
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyudaOficina 
         Height          =   300
         Left            =   1545
         TabIndex        =   7
         Top             =   2040
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   529
         XcodMaxLongitud =   0
         xcodwith        =   80
         NomTabla        =   "cp_oficina"
         TituloAyuda     =   "Ayuda de Tipo Analitico"
         ListaCampos     =   "vendedorcodigo(1),vendedornombres(1)"
         XcodCampo       =   "vendedorcodigo"
         XListCampo      =   "vendedornombres"
         ListaCamposDescrip=   "Código,Descripción"
         ListaCamposText =   "vendedorcodigo,vendedornombres"
         Requerido       =   0   'False
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyudaEmp 
         Height          =   315
         Left            =   1545
         TabIndex        =   2
         Top             =   270
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   556
         XcodMaxLongitud =   2
         xcodwith        =   80
         NomTabla        =   "co_multiempresas"
         TituloAyuda     =   "Busqueda de Empresas"
         ListaCampos     =   "empresacodigo(1),empresadescripcion(1)"
         XcodCampo       =   "empresacodigo"
         XListCampo      =   "empresadescripcion"
         ListaCamposDescrip=   "Código,Descripción"
         ListaCamposText =   "empresacodigo,empresadescripcion"
         Requerido       =   0   'False
      End
      Begin VB.Label Label3 
         Caption         =   "Empresa :"
         Height          =   240
         Left            =   150
         TabIndex        =   32
         Top             =   315
         Width           =   1005
      End
      Begin VB.Label Leoficina 
         Caption         =   "Oficina :"
         Height          =   255
         Left            =   150
         TabIndex        =   24
         Top             =   2055
         Width           =   840
      End
      Begin VB.Label lbccosto 
         AutoSize        =   -1  'True
         Caption         =   "C.Costo"
         Height          =   195
         Left            =   150
         TabIndex        =   23
         Top             =   1725
         Width           =   555
      End
      Begin VB.Label Lblanalitico 
         AutoSize        =   -1  'True
         Caption         =   "Entidad"
         Height          =   195
         Left            =   150
         TabIndex        =   20
         Top             =   1320
         Width           =   540
      End
      Begin VB.Label Le_Proveedor 
         Caption         =   "Proveedor :"
         Height          =   255
         Left            =   150
         TabIndex        =   19
         Top             =   1005
         Width           =   1260
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta deGastos:"
         Height          =   195
         Left            =   150
         TabIndex        =   8
         Top             =   675
         Width           =   1290
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Listar  Por"
      Height          =   1455
      Left            =   240
      TabIndex        =   0
      Top             =   2805
      Width           =   4635
      Begin VB.OptionButton Optmeses 
         Caption         =   "Resumido x Meses"
         Height          =   495
         Left            =   2880
         TabIndex        =   25
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton OptOficina 
         Caption         =   "Resumen x Oficina"
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   720
         Width           =   1335
      End
      Begin VB.OptionButton OptCCostos 
         Caption         =   "Resumen x C. Costos"
         Height          =   375
         Left            =   1440
         TabIndex        =   21
         Top             =   720
         Width           =   1695
      End
      Begin VB.OptionButton Optdetalle 
         Caption         =   "Detallado"
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton Optresumen 
         Caption         =   "Resumido General"
         Height          =   375
         Left            =   1440
         TabIndex        =   17
         Top             =   240
         Width           =   1575
      End
   End
End
Attribute VB_Name = "Frmreplistgastos"
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
Private Sub cmdaceptar_Click()
    Call imprimir
End Sub
Private Sub CmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call Ctr_Ayugastos.Conexion(VGCNx)
    Call Ctr_AyuProveedor.Conexion(VGCNx)
    Call Ctr_AyuEntidad.Conexion(VGcnxCT)
    Call Ctr_AyuCcosto.Conexion(VGcnxCT)
    Call Ctr_AyudaOficina.Conexion(VGCNx)
    Call Ctr_AyudaEmp.Conexion(VGCNx)
    Ctr_AyuCcosto.Filtro = "centrocostotipo=" & VGnumnivcos & " and centrocostocodigo<>'00' "
    Optresumen.Value = 1
    Optionproveedor.Value = 1
    DTPFechaIni = Date
    DTPFechaFin = Date
End Sub
Public Sub imprimir()
Dim arrform(3) As Variant, arrparm(12) As Variant
On Error GoTo Imprime
    Screen.MousePointer = 11
    '@BaseCompra, @BaseConta, @Prove, @Ano, @flagfecha, @Fechaini, @fechafin, @cuenta
    Dim rsdate As New ADODB.Recordset
    Set VGvardllgen = New dllgeneral.dll_general
    
    arrform(0) = "incluye=' '"
    arrform(1) = "xfechaini='" & DTPFechaIni.Value & "'"
    arrform(2) = "xfechafin='" & DTPFechaFin.Value & "'"
    
    arrparm(0) = VGParamSistem.BDEmpresa
    arrparm(1) = VGParamSistem.Anoproceso
    arrparm(2) = DTPFechaIni.Value
    arrparm(3) = DTPFechaFin.Value
    arrparm(4) = 1
 '   If Optdetalle.Value = True Then
        If Optionproveedor.Value = True Then
           arrparm(4) = 1
         Else
           arrparm(4) = 0
        End If
 '   End If
    arrparm(5) = "%%"
    arrparm(6) = "%%"
    arrparm(7) = "%%"
    arrparm(8) = "%%"
    arrparm(9) = "%%"
    arrparm(10) = 0
    arrparm(11) = "%%"
    If Chkdetraccion.Value = 1 Then
       arrparm(10) = 1
       arrform(0) = "incluye=' Solo detraccion '"
    End If
    
    If Trim(Ctr_Ayugastos.xclave) <> "" Then
    arrparm(5) = Ctr_Ayugastos.xclave
    End If
    If Trim(Ctr_AyuProveedor.xclave) <> "" Then
       arrparm(6) = Ctr_AyuProveedor.xclave
    End If
    If Trim(Ctr_AyuEntidad.xclave) <> "" Then
       arrparm(7) = Ctr_AyuEntidad.xclave
    End If
    If Trim(Ctr_AyuCcosto.xclave) <> "" Then
       arrparm(8) = Ctr_AyuCcosto.xclave
    End If
    If Trim(Ctr_AyudaOficina.xclave) <> "" Then
       arrparm(9) = Ctr_AyudaOficina.xclave
    End If
    If Trim(Ctr_AyudaEmp.xclave) <> "" Then
       arrparm(11) = Ctr_AyudaEmp.xclave
    End If
    
    If Optresumen.Value Then
       Call ImpresionRptProc("co_gastosResumen.rpt", arrform, arrparm, , "Gastos Resumidos ")
     ElseIf OptCCostos.Value Then
            Call ImpresionRptProc("co_gastosResumenCCostos.rpt", arrform, arrparm, , "Gastos Resumidos ")
          ElseIf OptOficina.Value Then
                 Call ImpresionRptProc("co_gastosResumenxOficina.rpt", arrform, arrparm, , "Gastos Resumidos x Oficina ")
             ElseIf Optmeses.Value Then
                    Call ImpresionRptProc("co_gastosResumenxMeses.rpt", arrform, arrparm, , "Gastos Resumidos x Meses ")
                  Else
                    Call ImpresionRptProc("co_gastosDetallado.rpt", arrform, arrparm, , "Gastos Detallados ")
    End If
    Screen.MousePointer = 1
    Exit Sub
Imprime:
Screen.MousePointer = 1
MsgBox err.Description
End Sub


