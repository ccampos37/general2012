VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "CONTROLAYUDA.OCX"
Begin VB.Form FrmRepHorExt 
   Caption         =   "Resumen de Horas Extras"
   ClientHeight    =   1365
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5385
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   1365
   ScaleWidth      =   5385
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   300
      Left            =   4110
      TabIndex        =   7
      Top             =   570
      Width           =   1095
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   300
      Left            =   4110
      TabIndex        =   6
      Top             =   270
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   1275
      Left            =   45
      TabIndex        =   0
      Top             =   30
      Width           =   5295
      Begin MSComCtl2.DTPicker DTPini 
         Height          =   315
         Left            =   1260
         TabIndex        =   5
         Top             =   270
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "MMMM - yyyy"
         Format          =   24641539
         CurrentDate     =   37648
      End
      Begin ctrlayuda_f.Ctr_Ayuda CtrAyu_Periodo 
         Height          =   315
         Left            =   1245
         TabIndex        =   3
         Top             =   900
         Width           =   3930
         _ExtentX        =   6932
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
      Begin MSComCtl2.DTPicker DTPfin 
         Height          =   315
         Left            =   1260
         TabIndex        =   8
         Top             =   585
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "MMMM - yyyy"
         Format          =   24641539
         CurrentDate     =   37648
      End
      Begin VB.Label Label3 
         Caption         =   "Tipo Trab :"
         Height          =   225
         Left            =   105
         TabIndex        =   4
         Top             =   975
         Width           =   885
      End
      Begin VB.Label Label2 
         Caption         =   "Periodo Fin :"
         Height          =   225
         Left            =   120
         TabIndex        =   2
         Top             =   660
         Width           =   1365
      End
      Begin VB.Label Label1 
         Caption         =   "Periodo Inicio : "
         Height          =   225
         Left            =   135
         TabIndex        =   1
         Top             =   330
         Width           =   1365
      End
   End
End
Attribute VB_Name = "FrmRepHorExt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdImprimir_Click()
    Call IMPRIMIR
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Width = 5505
    Me.Height = 1770
    Call CtrAyu_Periodo.Conexion(DBSYSTEM)
End Sub
Private Sub IMPRIMIR()
On Error GoTo imprime
    Dim arrform(2) As Variant, arrparm(8) As Variant, interval As Long
    '@Base,@FechaIni,@Intervalo,@Tipotrab
    DTPini.Day = 1: DTPfin.Day = 1
    interval = DateDiff("m", DTPini, DTPfin)
    arrparm(0) = REGSISTEMA.BASESQL
    arrparm(1) = FechS(DTPini, Sqlf)
    arrparm(2) = interval
    arrparm(3) = IIf(Trim(CtrAyu_Periodo.xClave) = "", "%%", CtrAyu_Periodo.xClave)
    Call ImpresionRptProc("pl_horextras.rpt", arrform, arrparm, , "Resumen de Horas Extras")
    Exit Sub
imprime:
    MsgBox ERR.Description
End Sub
