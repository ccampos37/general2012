VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Begin VB.Form FrmReporteOrdFabricacion 
   Caption         =   "Reporte de orden de fabricación"
   ClientHeight    =   2640
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8190
   LinkTopic       =   "Form1"
   ScaleHeight     =   2640
   ScaleWidth      =   8190
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   792
      Left            =   120
      TabIndex        =   9
      Top             =   1200
      Width           =   7155
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuordfab 
         Height          =   315
         Left            =   1800
         TabIndex        =   3
         Top             =   240
         Width           =   5220
         _ExtentX        =   9208
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
         Caption         =   "Orden de fabricacion:"
         Height          =   315
         Left            =   105
         TabIndex        =   10
         Top             =   315
         Width           =   1770
      End
   End
   Begin VB.Frame Frame2 
      Height          =   765
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   6240
      Begin VB.CheckBox ChkFech 
         Caption         =   "Rango de Fechas"
         Height          =   285
         Left            =   75
         TabIndex        =   0
         Top             =   -45
         Width           =   1620
      End
      Begin MSComCtl2.DTPicker DTPFechaIni 
         Height          =   300
         Left            =   1260
         TabIndex        =   1
         Top             =   315
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   52559873
         CurrentDate     =   37623.1285069444
      End
      Begin MSComCtl2.DTPicker DTPFechaFin 
         Height          =   300
         Left            =   4140
         TabIndex        =   2
         Top             =   315
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   52559873
         CurrentDate     =   37623.1264351852
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Inicio :"
         Height          =   210
         Left            =   150
         TabIndex        =   8
         Top             =   360
         Width           =   1110
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha Fin :"
         Height          =   210
         Left            =   3195
         TabIndex        =   7
         Top             =   375
         Width           =   810
      End
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   330
      Left            =   6480
      TabIndex        =   4
      Top             =   420
      Width           =   1260
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   330
      Left            =   6495
      TabIndex        =   6
      Top             =   810
      Width           =   1260
   End
End
Attribute VB_Name = "FrmReporteOrdFabricacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ChkFech_Click()
If ChkFech.Value = 1 Then
   DTPFechaIni.Enabled = True
   DTPFechaFin.Enabled = True
   Ctr_Ayuordfab.Enabled = False
 Else
   DTPFechaIni.Enabled = False
   DTPFechaFin.Enabled = False
   Ctr_Ayuordfab.Enabled = True
 End If
End Sub

Private Sub CmdAceptar_Click()
Dim arrform(1) As Variant, arrparm(5) As Variant
On Error GoTo Imprime
    Screen.MousePointer = 11
    '@BaseCompra, @BaseConta, @Prove, @Ano, @flagfecha, @Fechaini, @fechafin, @cuenta
    Dim rsdate As New ADODB.Recordset
    Set VGvardllgen = New dllgeneral.dll_general
    
    arrform(0) = "Rango=' Desde : " & DTPFechaIni.Value & "  Hasta " & DTPFechaFin.Value & "'"
    
    arrparm(0) = VGParamSistem.BDEmpresa
    arrparm(1) = ChkFech.Value
    arrparm(2) = IIf(Trim(Ctr_Ayuordfab.xclave) = "", "%%", Trim(Ctr_Ayuordfab.xclave))
    arrparm(3) = DTPFechaIni.Value
    arrparm(4) = DTPFechaFin.Value
    Call ImpresionRptProc("al_RelacionOrdendefabricacion.rpt", arrform, arrparm, , "Ordenes de fabricacion ")
    Screen.MousePointer = 1
    Exit Sub
Imprime:
Screen.MousePointer = 1
MsgBox Err.Description

End Sub

Private Sub CmdCancelar_Click()
Unload Me
End Sub

Private Sub Form_Load()

DTPFechaIni = Date
DTPFechaFin = Date

End Sub
