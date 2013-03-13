VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Begin VB.Form Frmreplistgastos 
   Caption         =   "Reporte de Gastos"
   ClientHeight    =   4845
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8190
   LinkTopic       =   "Form1"
   ScaleHeight     =   4845
   ScaleWidth      =   8190
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame5 
      Height          =   735
      Left            =   120
      TabIndex        =   18
      Top             =   120
      Width           =   7695
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuempresa 
         Height          =   315
         Left            =   1800
         TabIndex        =   19
         Top             =   240
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   556
         XcodMaxLongitud =   3
         xcodwith        =   300
         NomTabla        =   "co_multiempresas"
         TituloAyuda     =   "Busqueda de Empresas"
         ListaCampos     =   "empresacodigo(1),empresadescripcion(1)"
         XcodCampo       =   "empresacodigo"
         XListCampo      =   "empresadescripcion"
         ListaCamposDescrip=   "Codigo,Descripcion"
         ListaCamposText =   "empresacodigo,empresadescripcion"
         Requerido       =   0   'False
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Empresa"
         Height          =   195
         Index           =   1
         Left            =   480
         TabIndex        =   20
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame Frame4 
      Height          =   792
      Left            =   120
      TabIndex        =   14
      Top             =   960
      Width           =   7635
      Begin ctrlayuda_f.Ctr_Ayuda CtrAyu_caja 
         Height          =   315
         Left            =   1560
         TabIndex        =   15
         Top             =   240
         Width           =   6060
         _ExtentX        =   10689
         _ExtentY        =   556
         XcodMaxLongitud =   20
         xcodwith        =   1000
         NomTabla        =   "te_codigocaja"
         TituloAyuda     =   "Busqueda de Codigo de Caja"
         ListaCampos     =   "cajacodigo(1),cajadescripcion(1)"
         XcodCampo       =   "cajacodigo"
         XListCampo      =   "cajadescripcion"
         ListaCamposDescrip=   "Código,Descripción"
         ListaCamposText =   "cajacodigo,cajadescripcion"
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Codigo de caja"
         Height          =   195
         Index           =   0
         Left            =   105
         TabIndex        =   16
         Top             =   315
         Width           =   1065
      End
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   330
      Left            =   6600
      TabIndex        =   11
      Top             =   4320
      Width           =   1260
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   330
      Left            =   6600
      TabIndex        =   10
      Top             =   3840
      Width           =   1260
   End
   Begin VB.Frame Frame2 
      Height          =   765
      Left            =   240
      TabIndex        =   4
      Top             =   3840
      Width           =   6240
      Begin VB.CheckBox ChkFech 
         Caption         =   "Rango de Fechas"
         Height          =   285
         Left            =   75
         TabIndex        =   5
         Top             =   -45
         Width           =   1620
      End
      Begin MSComCtl2.DTPicker DTPFechaIni 
         Height          =   300
         Left            =   1260
         TabIndex        =   6
         Top             =   315
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   16777217
         CurrentDate     =   37623.1285069444
      End
      Begin MSComCtl2.DTPicker DTPFechaFin 
         Height          =   300
         Left            =   4140
         TabIndex        =   7
         Top             =   315
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   16777217
         CurrentDate     =   37623.1264351852
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha Fin :"
         Height          =   210
         Left            =   3195
         TabIndex        =   9
         Top             =   375
         Width           =   810
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Inicio :"
         Height          =   210
         Left            =   150
         TabIndex        =   8
         Top             =   360
         Width           =   1110
      End
   End
   Begin VB.Frame Frame1 
      Height          =   792
      Left            =   120
      TabIndex        =   1
      Top             =   1920
      Width           =   7635
      Begin ctrlayuda_f.Ctr_Ayuda CtrAyu_gastos 
         Height          =   312
         Left            =   1560
         TabIndex        =   2
         Top             =   252
         Width           =   6060
         _ExtentX        =   10689
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
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta deGastos:"
         Height          =   192
         Left            =   108
         TabIndex        =   3
         Top             =   312
         Width           =   1284
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Listar  Por"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   2880
      Width           =   7635
      Begin VB.OptionButton Optresumen 
         Caption         =   "Resumiido x Oficina"
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   17
         Top             =   240
         Width           =   2295
      End
      Begin VB.OptionButton Optdetalle 
         Caption         =   "Detallado"
         Height          =   375
         Left            =   5880
         TabIndex        =   13
         Top             =   240
         Width           =   972
      End
      Begin VB.OptionButton Optresumen 
         Caption         =   "Resumido"
         Height          =   375
         Index           =   0
         Left            =   3240
         TabIndex        =   12
         Top             =   240
         Width           =   1572
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
    DTPfechaini.Enabled = True
    DTPFechaFin.Enabled = True
  Else
    DTPfechaini.Enabled = False
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
    Call CtrAyu_gastos.Conexion(VGCNx)
    Call CtrAyu_caja.Conexion(VGCNx)
    Call Ctr_Ayuempresa.Conexion(VGCNx)
    Optresumen(0).Value = 1
    DTPfechaini = Date
    DTPFechaFin = Date
End Sub
Public Sub imprimir()
Dim arrform(2) As Variant, arrparm(7) As Variant
On Error GoTo Imprime
    Screen.MousePointer = 11
    '@BaseCompra, @BaseConta, @Prove, @Ano, @flagfecha, @Fechaini, @fechafin, @cuenta
    Dim rsdate As New ADODB.Recordset
    Set VGvardllgen = New dllgeneral.dll_general
    
    arrform(0) = "xfechaini='" & Format(DTPfechaini.Value, "dd/mm/yyyy") & "'"
    arrform(1) = "xfechafin='" & Format(DTPFechaFin.Value, "dd/mm/yyyy") & "'"
    
    arrparm(0) = VGParamSistem.BDEmpresa
    arrparm(1) = IIf(Trim(CtrAyu_caja.xclave) = "", "%%", Trim(CtrAyu_caja.xclave))
    arrparm(2) = IIf(Trim(CtrAyu_gastos.xclave) = "", "%%", Trim(CtrAyu_gastos.xclave))
    arrparm(3) = DTPfechaini.Value
    arrparm(4) = DTPFechaFin.Value
    arrparm(5) = "C"
    arrparm(6) = IIf(Ctr_Ayuempresa.xclave = Empty, "%%", Ctr_Ayuempresa.xclave)

    If Optresumen(0).Value Then
       Call ImpresionRptProc("te_gastosResumen.rpt", arrform, arrparm, , "Gastos Resumidos ")
     ElseIf Optresumen(1).Value Then
            Call ImpresionRptProc("te_gastosResumen_oficina.rpt", arrform, arrparm, , "Gastos Resumidos ")
         Else
         Call ImpresionRptProc("te_gastosDetallado.rpt", arrform, arrparm, , "Gastos Detallados ")
    End If
    Screen.MousePointer = 1
    Exit Sub
Imprime:
Screen.MousePointer = 1
MsgBox Err.Description
End Sub


