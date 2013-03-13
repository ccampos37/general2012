VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "CONTROLAYUDA.OCX"
Begin VB.Form FrmRepCtahistprog 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Historial de Ctas. Ctes. Programadas Pendientes"
   ClientHeight    =   3045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6315
   Icon            =   "FrmRepCtahistprog.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   6315
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   765
      Left            =   90
      TabIndex        =   10
      Top             =   2145
      Width           =   6120
      Begin VB.CheckBox ChkFech 
         Caption         =   "Rango de Fechas"
         Height          =   285
         Left            =   75
         TabIndex        =   14
         Top             =   -45
         Width           =   1620
      End
      Begin MSComCtl2.DTPicker DTPFechaIni 
         Height          =   300
         Left            =   1260
         TabIndex        =   13
         Top             =   315
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   24510465
         CurrentDate     =   37623.1285069444
      End
      Begin MSComCtl2.DTPicker DTPFechaFin 
         Height          =   300
         Left            =   4140
         TabIndex        =   15
         Top             =   315
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   24510465
         CurrentDate     =   37623.1264351852
      End
      Begin VB.Label Label5 
         Caption         =   "Fecha Inicio :"
         Height          =   210
         Left            =   150
         TabIndex        =   17
         Top             =   360
         Width           =   1110
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Fin :"
         Height          =   210
         Left            =   3195
         TabIndex        =   16
         Top             =   375
         Width           =   810
      End
   End
   Begin ctrlayuda_f.Ctr_Ayuda CtrAyu_trab 
      Height          =   315
      Left            =   1125
      TabIndex        =   7
      Top             =   1050
      Width           =   5070
      _ExtentX        =   8943
      _ExtentY        =   556
      XcodMaxLongitud =   6
      xcodwith        =   400
      NomTabla        =   "VWTRABAJGEN"
      TituloAyuda     =   "Busqueda del Trabajador"
      ListaCampos     =   "CODTRAB(1),NOMBRES(1)"
      XcodCampo       =   "CODTRAB"
      XListCampo      =   "NOMBRES"
      ListaCamposDescrip=   "Código,Nombres"
      ListaCamposText =   "CODTRAB,NOMBRES"
      Requerido       =   0   'False
   End
   Begin VB.Frame Frame1 
      Caption         =   "Opciones"
      Height          =   825
      Left            =   90
      TabIndex        =   2
      Top             =   45
      Width           =   4200
      Begin VB.OptionButton xTodos 
         Caption         =   "Todos"
         Height          =   300
         Left            =   2910
         TabIndex        =   5
         Top             =   315
         Width           =   1050
      End
      Begin VB.OptionButton xEgresos 
         Caption         =   "&Egresos"
         Height          =   210
         Left            =   1575
         TabIndex        =   4
         Top             =   375
         Width           =   1050
      End
      Begin VB.OptionButton XIngresos 
         Caption         =   "&Ingresos"
         Height          =   240
         Left            =   120
         TabIndex        =   3
         Top             =   345
         Width           =   1050
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancelar"
      Height          =   330
      Left            =   4995
      TabIndex        =   1
      Top             =   540
      Width           =   1125
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   345
      Left            =   4995
      TabIndex        =   0
      Top             =   150
      Width           =   1140
   End
   Begin ctrlayuda_f.Ctr_Ayuda CtrAyu_Periodo 
      Height          =   315
      Left            =   1125
      TabIndex        =   9
      Top             =   1695
      Width           =   5070
      _ExtentX        =   8943
      _ExtentY        =   556
      XcodMaxLongitud =   2
      NomTabla        =   "TIPOSTRAB"
      TituloAyuda     =   "Busqueda de Trabajador"
      ListaCampos     =   "TIPTRAB(1),DESCRIP(1)"
      XcodCampo       =   "TIPTRAB"
      XListCampo      =   "DESCRIP"
      ListaCamposDescrip=   "Código,Descripción"
      ListaCamposText =   "TIPTRAB,DESCRIP"
      Requerido       =   0   'False
   End
   Begin ctrlayuda_f.Ctr_Ayuda CtrAyu_Concepto 
      Height          =   315
      Left            =   1125
      TabIndex        =   8
      Top             =   1380
      Width           =   5070
      _ExtentX        =   8943
      _ExtentY        =   556
      XcodMaxLongitud =   10
      xcodwith        =   700
      NomTabla        =   "CTAGRUPO"
      TituloAyuda     =   "Busqueda de Concepto Cta. Cte."
      ListaCampos     =   "CODGRUPO(1),NOMBRE(1)"
      XcodCampo       =   "CODGRUPO"
      XListCampo      =   "NOMBRE"
      ListaCamposDescrip=   "Código,Descripción"
      ListaCamposText =   "CODGRUPO,NOMBRE"
      Requerido       =   0   'False
   End
   Begin VB.Label Label3 
      Caption         =   "Concepto :"
      Height          =   225
      Left            =   120
      TabIndex        =   12
      Top             =   1455
      Width           =   885
   End
   Begin VB.Label Label2 
      Caption         =   "Tipo Trab :"
      Height          =   225
      Left            =   120
      TabIndex        =   11
      Top             =   1770
      Width           =   885
   End
   Begin VB.Label Label1 
      Caption         =   "Trabajador :"
      Height          =   225
      Left            =   135
      TabIndex        =   6
      Top             =   1110
      Width           =   885
   End
End
Attribute VB_Name = "FrmRepCtahistprog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TIPO As String

Private Sub ChkFech_Click()
If ChkFech.Value = 1 Then
    DTPFechaIni.Enabled = True
    DTPFechaFin.Enabled = True
  Else
    DTPFechaIni.Enabled = False
    DTPFechaFin.Enabled = False
End If
End Sub

Private Sub Command1_Click()
    Call IMPRIMIR
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    Call CtrAyu_trab.Conexion(DBSYSTEM)
    Call CtrAyu_Periodo.Conexion(DBSYSTEM)
    Call CtrAyu_Concepto.Conexion(DBSYSTEM)
    xEgresos.Value = True
End Sub
Private Sub IMPRIMIR()
    Dim arrform(2) As Variant, arrparm(8) As Variant
    '@BASE, @CODTRAB, @GRUPO, @TIPO, @TipoTrab, @flagfecha, @Fechaini, @fechafin
    arrparm(0) = REGSISTEMA.BASESQL
    arrparm(1) = ESNULO(CtrAyu_trab.xClave, "%%")
    arrparm(2) = ESNULO(CtrAyu_Concepto.xClave, "%%")
    arrparm(3) = TIPO
    arrparm(4) = ESNULO(CtrAyu_Periodo.xClave, "%%") 'Tipo de Trabajador
    arrparm(5) = IIf(ChkFech.Value = 1, "0", "1")
    arrparm(6) = FechS(DTPFechaIni.Value, Sqlf)
    arrparm(7) = FechS(DTPFechaFin.Value, Sqlf)
    
    Call ImpresionRptProc("pl_ctacteprog.rpt", arrform, arrparm, , "Historial de Cta. Cte Trabajador - pl_ctactehist.rpt")

End Sub
    

Private Sub xEgresos_Click()
    TIPO = "2"
End Sub

Private Sub XIngresos_Click()
    TIPO = "1"
End Sub

Private Sub XTODOS_Click()
    TIPO = "%%"
End Sub
