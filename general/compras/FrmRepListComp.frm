VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Begin VB.Form FrmRepListCuenta 
   Caption         =   "Listado de Compra x Cuenta"
   ClientHeight    =   2868
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   7632
   Icon            =   "FrmRepListComp.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   2868
   ScaleWidth      =   7632
   Begin VB.Frame FrmChk 
      Height          =   735
      Left            =   60
      TabIndex        =   13
      Top             =   1110
      Width           =   7515
      Begin VB.CheckBox chkflagmodo 
         Caption         =   "Lista por "
         Height          =   285
         Left            =   75
         TabIndex        =   17
         Top             =   0
         Width           =   1020
      End
      Begin VB.CheckBox ChkActCaja 
         Alignment       =   1  'Right Justify
         Caption         =   "Caja Chica"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   5580
         TabIndex        =   16
         Top             =   360
         Width           =   1470
      End
      Begin VB.CheckBox ChkRegComp 
         Alignment       =   1  'Right Justify
         Caption         =   "Registro. Compra"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   2910
         TabIndex        =   15
         Top             =   360
         Width           =   1755
      End
      Begin VB.CheckBox ChkCtaCte 
         Alignment       =   1  'Right Justify
         Caption         =   "Cuenta Corriente"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   510
         TabIndex        =   14
         Top             =   360
         Width           =   1725
      End
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   330
      Left            =   6300
      TabIndex        =   9
      Top             =   2445
      Width           =   1260
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   330
      Left            =   6285
      TabIndex        =   8
      Top             =   2055
      Width           =   1260
   End
   Begin VB.Frame Frame2 
      Height          =   765
      Left            =   60
      TabIndex        =   4
      Top             =   2010
      Width           =   6120
      Begin MSComCtl2.DTPicker DTPFechaIni 
         Height          =   300
         Left            =   1260
         TabIndex        =   10
         Top             =   315
         Width           =   1785
         _ExtentX        =   3154
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   62783489
         CurrentDate     =   37623.1285069444
      End
      Begin VB.CheckBox ChkFech 
         Caption         =   "Rango de Fechas"
         Height          =   285
         Left            =   75
         TabIndex        =   5
         Top             =   -45
         Width           =   1620
      End
      Begin MSComCtl2.DTPicker DTPFechaFin 
         Height          =   300
         Left            =   4140
         TabIndex        =   11
         Top             =   315
         Width           =   1785
         _ExtentX        =   3154
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   62783489
         CurrentDate     =   37623.1264351852
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha Fin :"
         Height          =   210
         Left            =   3195
         TabIndex        =   7
         Top             =   375
         Width           =   810
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Inicio :"
         Height          =   210
         Left            =   150
         TabIndex        =   6
         Top             =   360
         Width           =   1110
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1155
      Left            =   45
      TabIndex        =   0
      Top             =   -45
      Width           =   7515
      Begin ctrlayuda_f.Ctr_Ayuda CtrAyu_Cuenta 
         Height          =   315
         Left            =   1485
         TabIndex        =   2
         Top             =   615
         Width           =   5940
         _ExtentX        =   10478
         _ExtentY        =   550
         XcodMaxLongitud =   20
         xcodwith        =   1000
         NomTabla        =   "ct_cuenta"
         TituloAyuda     =   "Busqueda de Cuenta"
         ListaCampos     =   $"FrmRepListComp.frx":1E72
         XcodCampo       =   "cuentacodigo"
         XListCampo      =   "cuentadescripcion"
         ListaCamposDescrip=   "Código,Descripción"
         ListaCamposText =   "cuentacodigo,cuentadescripcion,cuentaestadoccostos,cuentaestadoanalitico,cuentadocumento,tipoanaliticocodigo,tipoajuste"
         Requerido       =   0   'False
      End
      Begin ctrlayuda_f.Ctr_Ayuda CtrAyu_Modoprovi 
         Height          =   315
         Left            =   1485
         TabIndex        =   12
         Top             =   240
         Width           =   5940
         _ExtentX        =   10478
         _ExtentY        =   550
         XcodMaxLongitud =   2
         NomTabla        =   "co_modoprovi"
         TituloAyuda     =   "Busqueda de Modo de Compra"
         ListaCampos     =   "modoprovicod(1), modoprovidesc(1),modoprovictacte(3), modoproviregcom(3), modoprovitesor(3)"
         XcodCampo       =   "modoprovicod"
         XListCampo      =   "modoprovidesc"
         ListaCamposDescrip=   "Código,Descripción"
         ListaCamposText =   "modoprovicod, modoprovidesc,modoprovictacte, modoproviregcom, modoprovitesor"
         Requerido       =   0   'False
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta :"
         Height          =   195
         Left            =   105
         TabIndex        =   3
         Top             =   675
         Width           =   600
      End
      Begin VB.Label Le_Proveedor 
         Caption         =   "Modo de Compra :"
         Height          =   255
         Left            =   90
         TabIndex        =   1
         Top             =   270
         Width           =   1575
      End
   End
End
Attribute VB_Name = "FrmRepListCuenta"
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
    Me.Width = 7755
    Me.Height = 3270
    Call CtrAyu_Modoprovi.Conexion(VGcnx)
    Call Ctrayu_cuenta.Conexion(VGcnxCT)
End Sub
Public Sub imprimir()
Dim arrform(0) As Variant, arrparm(13) As Variant
On Error GoTo Imprime
    Screen.MousePointer = 11
    '@BaseCompra, @BaseConta, @Prove, @Ano, @flagfecha, @Fechaini, @fechafin, @cuenta
    Dim rsdate As New ADODB.Recordset
    Set VGvardllgen = New dllgeneral.dll_general
    arrparm(0) = VGParamSistem.BDEmpresa
    arrparm(1) = VGParamSistem.BDEmpresaCT
    arrparm(2) = VGParamSistem.BDEmpresaCP
    arrparm(3) = VGvardllgen.ESNULO(Trim(CtrAyu_Modoprovi.xclave), "%%")
    arrparm(4) = VGParamSistem.Anoproceso
    arrparm(5) = IIf(ChkFech.Value = 1, "0", "1")
    arrparm(6) = FechS(DTPFechaIni.Value, 1)
    arrparm(7) = FechS(DTPFechaFin.Value, 1)
    arrparm(8) = VGvardllgen.ESNULO(Trim(Ctrayu_cuenta.xclave), "%") & "%"
    arrparm(9) = IIf(chkflagmodo.Value = 1, "0", "1")
    arrparm(10) = IIf(ChkCtaCte.Value = 1, "1", "0")
    arrparm(11) = IIf(ChkRegComp.Value = 1, "1", "0")
    arrparm(12) = IIf(ChkActCaja.Value = 1, "1", "0")
    Call ImpresionRptProc("rptcolistacuenta.rpt", arrform, arrparm, , "Registro de Compras ")
    Screen.MousePointer = 1
    Exit Sub
Imprime:
Screen.MousePointer = 1
MsgBox Err.Description
End Sub

