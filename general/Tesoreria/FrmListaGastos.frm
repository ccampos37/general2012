VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Begin VB.Form FrmListaGastos 
   Caption         =   "Lista Gastos"
   ClientHeight    =   4032
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   9060
   LinkTopic       =   "Form1"
   ScaleHeight     =   4032
   ScaleWidth      =   9060
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Listar  Por"
      Height          =   735
      Left            =   0
      TabIndex        =   16
      Top             =   2040
      Width           =   7635
      Begin VB.CheckBox ChkDetallado 
         Alignment       =   1  'Right Justify
         Caption         =   "Detallado"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   2910
         TabIndex        =   18
         Top             =   360
         Width           =   1755
      End
      Begin VB.CheckBox ChkResumen 
         Alignment       =   1  'Right Justify
         Caption         =   "Resumido"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   510
         TabIndex        =   17
         Top             =   360
         Width           =   1725
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1155
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   7635
      Begin ctrlayuda_f.Ctr_Ayuda CtrAyu_gastos 
         Height          =   315
         Left            =   1485
         TabIndex        =   14
         Top             =   615
         Width           =   6060
         _ExtentX        =   10689
         _ExtentY        =   550
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
         Caption         =   "Cuenta :"
         Height          =   195
         Left            =   105
         TabIndex        =   15
         Top             =   675
         Width           =   600
      End
   End
   Begin VB.Frame Frame2 
      Height          =   765
      Left            =   15
      TabIndex        =   7
      Top             =   3015
      Width           =   6240
      Begin VB.CheckBox ChkFech 
         Caption         =   "Rango de Fechas"
         Height          =   285
         Left            =   75
         TabIndex        =   8
         Top             =   -45
         Width           =   1620
      End
      Begin MSComCtl2.DTPicker DTPFechaIni 
         Height          =   300
         Left            =   1260
         TabIndex        =   9
         Top             =   315
         Width           =   1785
         _ExtentX        =   3154
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   23658497
         CurrentDate     =   37623.1285069444
      End
      Begin MSComCtl2.DTPicker DTPFechaFin 
         Height          =   300
         Left            =   4140
         TabIndex        =   10
         Top             =   315
         Width           =   1785
         _ExtentX        =   3154
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   23658497
         CurrentDate     =   37623.1264351852
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Inicio :"
         Height          =   210
         Left            =   150
         TabIndex        =   12
         Top             =   360
         Width           =   1110
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha Fin :"
         Height          =   210
         Left            =   3195
         TabIndex        =   11
         Top             =   375
         Width           =   810
      End
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   330
      Left            =   6360
      TabIndex        =   6
      Top             =   3060
      Width           =   1260
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   330
      Left            =   6375
      TabIndex        =   5
      Top             =   3450
      Width           =   1260
   End
   Begin VB.Frame FrmChk 
      Height          =   735
      Left            =   15
      TabIndex        =   0
      Top             =   1155
      Width           =   7635
      Begin VB.CheckBox ChkCtaCte 
         Alignment       =   1  'Right Justify
         Caption         =   "Cuenta Corriente"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   510
         TabIndex        =   4
         Top             =   360
         Width           =   1725
      End
      Begin VB.CheckBox ChkRegComp 
         Alignment       =   1  'Right Justify
         Caption         =   "Registro. Compra"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   2910
         TabIndex        =   3
         Top             =   360
         Width           =   1755
      End
      Begin VB.CheckBox ChkActCaja 
         Alignment       =   1  'Right Justify
         Caption         =   "Caja Chica"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   5580
         TabIndex        =   2
         Top             =   360
         Width           =   1470
      End
      Begin VB.CheckBox chkflagmodo 
         Caption         =   "Lista por "
         Height          =   285
         Left            =   75
         TabIndex        =   1
         Top             =   0
         Width           =   1020
      End
   End
End
Attribute VB_Name = "FrmListaGastos"
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

Private Sub cmdAceptar_Click()
    Call Imprimir
End Sub
Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub CtrAyu_Cuenta_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)

End Sub

Private Sub Form_Load()
'    Call CtrAyu_Modoprovi.Conexion(VGcnx)
    Call CtrAyu_gastos.Conexion(Cn)
    ChkResumen.Value = 1
    DTPfechaini = Date
    DTPFechaFin = Date
End Sub
Public Sub Imprimir()
Dim arrform(0) As Variant, arrparm(10) As Variant
On Error GoTo Imprime
    Screen.MousePointer = 11
    '@BaseCompra, @BaseConta, @Prove, @Ano, @flagfecha, @Fechaini, @fechafin, @cuenta
    Dim rsdate As New ADODB.Recordset
    Set VGvardllgen = New dllgeneral.dll_general
    arrparm(0) = Cn.DefaultDatabase
    arrparm(1) = cnconta.DefaultDatabase
    arrparm(2) = Cn.DefaultDatabase
    arrparm(3) = "%%"
    arrparm(5) = IIf(ChkFech.Value = 1, "0", "1")
    arrparm(6) = FechS(DTPfechaini.Value, 1)
    arrparm(7) = FechS(DTPFechaFin.Value, 1)
    arrparm(8) = VGvardllgen.ESNULO(Trim(CtrAyu_gastos.xclave), "%") & "%"
    arrparm(9) = IIf(chkflagmodo.Value = 1, "0", "1")
'   arrparm(10) = IIf(ChkCtaCte.Value = 1, "1", "0")
'   arrparm(11) = IIf(ChkRegComp.Value = 1, "1", "0")
'   arrparm(12) = IIf(ChkActCaja.Value = 1, "1", "0")
    If optResumen = True Then
       Call ImpresionRptProc2("te_gastosResumen.rpt", arrform, arrparm, , "Gastos Resumidos ")
     Else
       Call ImpresionRptProc2("te_gastosDetallado.rpt", arrform, arrparm, , "Gastos Detallados ")
    End If
    Screen.MousePointer = 1
    Exit Sub
Imprime:
Screen.MousePointer = 1
MsgBox Err.Description
End Sub



