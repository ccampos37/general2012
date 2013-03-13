VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Begin VB.Form RptModoVenta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Modo de Venta"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6450
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   6450
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   2265
      Left            =   105
      TabIndex        =   2
      Top             =   60
      Width           =   6180
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_PuntoVta 
         Height          =   330
         Left            =   1665
         TabIndex        =   9
         Top             =   1125
         Width           =   4305
         _ExtentX        =   7594
         _ExtentY        =   582
         XcodMaxLongitud =   0
         xcodwith        =   800
         NomTabla        =   "vt_puntoventa"
         ListaCampos     =   "puntovtacodigo(1),puntovtadescripcion(1)"
         XcodCampo       =   "puntovtacodigo"
         XListCampo      =   "puntovtadescripcion"
         ListaCamposDescrip=   "Código,Descripción"
         ListaCamposText =   "puntovtacodigo,puntovtadescripcion"
         Requerido       =   0   'False
      End
      Begin MSComCtl2.DTPicker DTP_FechaInicio 
         Height          =   330
         Left            =   1665
         TabIndex        =   3
         Top             =   360
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         _Version        =   393216
         Format          =   117768193
         CurrentDate     =   37586
      End
      Begin MSComCtl2.DTPicker DTP_FechaFin 
         Height          =   330
         Left            =   1665
         TabIndex        =   4
         Top             =   735
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         _Version        =   393216
         Format          =   117768193
         CurrentDate     =   37586
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_ModoVta 
         Height          =   330
         Left            =   1665
         TabIndex        =   10
         Top             =   1530
         Width           =   4305
         _ExtentX        =   7594
         _ExtentY        =   582
         XcodMaxLongitud =   0
         xcodwith        =   800
         NomTabla        =   "vt_modoventa"
         ListaCampos     =   "modovtacodigo(1),modovtadescripcion(1)"
         XcodCampo       =   "modovtacodigo"
         XListCampo      =   "modovtadescripcion"
         ListaCamposDescrip=   "Código,Descripción"
         ListaCamposText =   "modovtacodigo,modovtadescripcion"
         Requerido       =   0   'False
      End
      Begin VB.Label Label1 
         Caption         =   "Modo de Venta"
         Height          =   255
         Index           =   5
         Left            =   390
         TabIndex        =   8
         Top             =   1560
         Width           =   1185
      End
      Begin VB.Label Label1 
         Caption         =   "Punto de Venta"
         Height          =   255
         Index           =   4
         Left            =   390
         TabIndex        =   7
         Top             =   1170
         Width           =   1245
      End
      Begin VB.Label Label1 
         Caption         =   "Hasta la Fecha"
         Height          =   255
         Index           =   2
         Left            =   390
         TabIndex        =   6
         Top             =   795
         Width           =   1245
      End
      Begin VB.Label Label1 
         Caption         =   "Desde la Fecha"
         Height          =   255
         Index           =   0
         Left            =   390
         TabIndex        =   5
         Top             =   420
         Width           =   1245
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   360
      Left            =   1793
      TabIndex        =   1
      Top             =   2610
      Width           =   1305
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   360
      Left            =   3413
      TabIndex        =   0
      Top             =   2610
      Width           =   1245
   End
End
Attribute VB_Name = "RptModoVenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adll As New dllgeneral.dll_general

Private Sub Form_Load()
   MostrarForm Me, "C2"
   DTP_FechaInicio.Value = "01/" & Format(Month(Now), "00") & "/" & Year(Date)
   DTP_FechaFin.Value = Format(Date, "dd/mm/yyyy")
   Ctr_ModoVta.conexion cn
   Ctr_PuntoVta.conexion cn
End Sub

Private Sub cmdAceptar_Click()
   Call Imprimir
End Sub

Private Sub cmdCancelar_Click()
   Unload Me
End Sub

Sub Imprimir()
'FIXIT: Declare 'arrform' con un tipo de datos de enlace en tiempo de compilación          FixIT90210ae-R1672-R1B8ZE
Dim arrform(4) As Variant, arrparm(5) As Variant
Dim NombreRep As String, CadOrden As String
Dim nombrerepSub As String
Dim mon As String
    arrparm(0) = VGCNx.DefaultDatabase
    arrparm(1) = IIf(Ctr_PuntoVta.xclave = Empty, "%%", RTrim$(Ctr_PuntoVta.xclave))
    arrparm(2) = IIf(Ctr_ModoVta.xclave = Empty, "%%", RTrim$(Ctr_ModoVta.xclave))
    arrparm(3) = DTP_FechaInicio.Value
    arrparm(4) = DTP_FechaFin.Value
    arrform(0) = "Desde='" & Format(DTP_FechaInicio.Value, "dd/mm/yyyy") & "'"
    arrform(1) = "Hasta='" & Format(DTP_FechaFin.Value, "dd/mm/yyyy") & "'"
    arrform(2) = "Puntoventa='" & IIf(Ctr_PuntoVta.xclave = Empty, "Todos", RTrim$(Ctr_PuntoVta.xclave)) & "'"
    arrform(3) = "Modoventa='" & IIf(Ctr_ModoVta.xclave = Empty, "Todos", RTrim$(Ctr_ModoVta.xclave)) & "'"
    NombreRep = "RepvtVtasxModoVta.rpt"
    nombrerepSub = "RepvtSubVtasxModoVta.rpt"

    CadOrden = ""
    Call ImpresionRpt_SubRpt_Proc(NombreRep, arrform, arrparm, nombrerepSub, CadOrden, "Cuenta Corriente por Cliente")
End Sub

