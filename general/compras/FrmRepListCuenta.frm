VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Begin VB.Form FrmReplistcuenta 
   Caption         =   "Form1"
   ClientHeight    =   4290
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6855
   LinkTopic       =   "Form1"
   ScaleHeight     =   4290
   ScaleWidth      =   6855
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   1560
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   6660
      Begin ctrlayuda_f.Ctr_Ayuda CtrAyu_Cuenta 
         Height          =   330
         Left            =   1657
         TabIndex        =   12
         Top             =   1125
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   582
         XcodMaxLongitud =   20
         xcodwith        =   1000
         NomTabla        =   "ct_cuenta"
         TituloAyuda     =   "Busqueda de Cuenta"
         ListaCampos     =   $"FrmReplistcuenta.frx":0000
         XcodCampo       =   "cuentacodigo"
         XListCampo      =   "cuentadescripcion"
         ListaCamposDescrip=   "Código,Descripción"
         ListaCamposText =   "cuentacodigo,cuentadescripcion,cuentaestadoccostos,cuentaestadoanalitico,cuentadocumento,tipoanaliticocodigo,tipoajuste"
         Requerido       =   0   'False
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyuTipoCompra 
         Height          =   330
         Left            =   1657
         TabIndex        =   13
         Top             =   720
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   582
         XcodMaxLongitud =   2
         NomTabla        =   "co_tipocompra"
         TituloAyuda     =   "Busqueda de Tipo de Compra"
         ListaCampos     =   "tipocompracodigo(1), tipocompradesc(1),tipocomprainafecta(1)"
         XcodCampo       =   "tipocompracodigo"
         XListCampo      =   "tipocompradesc"
         ListaCamposDescrip=   "Código,Descripción"
         ListaCamposText =   "tipocompracodigo, tipocompradesc,tipocomprainafecta"
         Requerido       =   0   'False
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuempresa 
         Height          =   330
         Left            =   1657
         TabIndex        =   14
         Top             =   225
         Width           =   4740
         _ExtentX        =   8361
         _ExtentY        =   582
         XcodMaxLongitud =   3
         xcodwith        =   300
         NomTabla        =   "co_multiempresas"
         TituloAyuda     =   "Busqueda de Empresas"
         ListaCampos     =   "empresacodigo(1),empresadescripcion(1),agentederetencion(1)"
         XcodCampo       =   "empresacodigo"
         XListCampo      =   "empresadescripcion"
         ListaCamposDescrip=   "Codigo,Descripcion"
         ListaCamposText =   "empresacodigo,empresadescripcion,agentederetencion"
      End
      Begin VB.Label Le_Proveedor 
         Caption         =   "Tipo de Compra :"
         Height          =   255
         Left            =   270
         TabIndex        =   17
         Top             =   765
         Width           =   1440
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta :"
         Height          =   195
         Left            =   285
         TabIndex        =   16
         Top             =   1170
         Width           =   600
      End
      Begin VB.Label Leempresa 
         AutoSize        =   -1  'True
         Caption         =   "Empresa :"
         Height          =   195
         Left            =   315
         TabIndex        =   15
         Top             =   285
         Width           =   705
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1125
      Left            =   90
      TabIndex        =   5
      Top             =   2415
      Width           =   5040
      Begin VB.CheckBox ChkFech 
         Caption         =   "Rango de Fechas"
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   75
         TabIndex        =   7
         Top             =   -45
         Width           =   1620
      End
      Begin MSComCtl2.DTPicker DTPFechaIni 
         Height          =   300
         Left            =   540
         TabIndex        =   6
         Top             =   630
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   57409537
         CurrentDate     =   37623.1285069444
      End
      Begin MSComCtl2.DTPicker DTPFechaFin 
         Height          =   300
         Left            =   2970
         TabIndex        =   8
         Top             =   630
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   57409537
         CurrentDate     =   37623.1264351852
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Inicio :"
         Height          =   210
         Left            =   870
         TabIndex        =   10
         Top             =   270
         Width           =   1110
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha Fin :"
         Height          =   210
         Left            =   3420
         TabIndex        =   9
         Top             =   240
         Width           =   810
      End
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   330
      Left            =   5295
      TabIndex        =   4
      Top             =   2670
      Width           =   1260
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   330
      Left            =   5310
      TabIndex        =   3
      Top             =   3075
      Width           =   1260
   End
   Begin VB.Frame Frame3 
      Caption         =   "Listar Por "
      ForeColor       =   &H00FF8080&
      Height          =   735
      Left            =   75
      TabIndex        =   0
      Top             =   1605
      Width           =   6510
      Begin VB.OptionButton OptDetallado 
         Caption         =   "Detallado"
         Height          =   495
         Left            =   1080
         TabIndex        =   2
         Top             =   120
         Width           =   1215
      End
      Begin VB.OptionButton OptResumido 
         Caption         =   "Resumido"
         Height          =   495
         Left            =   3120
         TabIndex        =   1
         Top             =   120
         Width           =   1215
      End
   End
End
Attribute VB_Name = "FrmReplistcuenta"
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
Private Sub Form_Load()
    Dim cfecha As Date
    Call Ctr_AyuTipoCompra.Conexion(VGCNx)
    Call CtrAyu_Cuenta.Conexion(VGcnxCT)
    Call Ctr_Ayuempresa.Conexion(VGCNx)
    DTPFechaIni.Value = Format("01/" & Format(Month(VGfecha), "00") & "/" & Year(VGfecha), "dd/mm/yyyy")
    cfecha = Format("01/" & Format(Month(VGfecha) + 1, "00") & "/" & Year(VGfecha), "dd/mm/yyyy")
    DTPFechaFin.Value = cfecha - 1

End Sub
Public Sub imprimir()
Dim arrform(2) As Variant, arrparm(9) As Variant
Dim Reporte As String
On Error GoTo Imprime
    Screen.MousePointer = 11
    '@BaseCompra, @BaseConta, @Prove, @Ano, @flagfecha, @Fechaini, @fechafin, @cuenta
    Dim rsdate As New ADODB.Recordset
    Set VGvardllgen = New dllgeneral.dll_general
    arrform(0) = "fechaini='" & DTPFechaIni.Value & "'"
    arrform(1) = "fechafin='" & DTPFechaFin.Value & "'"
    
    arrparm(0) = VGParamSistem.BDEmpresa
    arrparm(1) = VGParamSistem.BDEmpresaCT
    arrparm(2) = IIf(Ctr_Ayuempresa.xclave = "", "%%", Ctr_Ayuempresa.xclave)
    arrparm(3) = VGvardllgen.ESNULO(Trim(Ctr_AyuTipoCompra.xclave), "%%")
    arrparm(4) = VGParamSistem.Anoproceso
    arrparm(5) = IIf(ChkFech.Value = 1, "0", "1")
    arrparm(6) = DTPFechaIni.Value
    arrparm(7) = DTPFechaFin.Value
    arrparm(8) = VGvardllgen.ESNULO(Trim(CtrAyu_Cuenta.xclave), "%") & "%"
    If OptDetallado.Value Then
       Call ImpresionRptProc("co_listacuenta.rpt", arrform, arrparm, , "Registro de provisiones VS Ctas Contables ")
      Else
       Call ImpresionRptProc("co_listacuentaResumen.rpt", arrform, arrparm, , "Registro de Compras ")
    End If
    Screen.MousePointer = 1
    Exit Sub
Imprime:
Screen.MousePointer = 1
MsgBox Err.Description
End Sub

Private Sub Option1_Click()

End Sub

