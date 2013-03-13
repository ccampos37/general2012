VERSION 5.00
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Begin VB.Form frmRepListadoCtasDist 
   Caption         =   "Reporte de Cuentas Distribución"
   ClientHeight    =   2970
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5385
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   2970
   ScaleWidth      =   5385
   Begin VB.Frame Frame1 
      Caption         =   "Seleccionar Cuenta"
      Height          =   1560
      Left            =   0
      TabIndex        =   1
      Top             =   180
      Width           =   5355
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuda1 
         Height          =   555
         Left            =   90
         TabIndex        =   0
         Top             =   495
         Width           =   5205
         _ExtentX        =   9181
         _ExtentY        =   979
         XcodMaxLongitud =   0
         xcodwith        =   800
         NomTabla        =   "ct_cuenta"
         ListaCampos     =   "cuentacodigo(1),cuentadescripcion(1)"
         XcodCampo       =   "cuentacodigo"
         XListCampo      =   "cuentadescripcion"
         ListaCamposDescrip=   "Código,Descripción"
         ListaCamposText =   "cuentacodigo,cuentadescripcion"
      End
   End
End
Attribute VB_Name = "frmRepListadoCtasDist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
  Ctr_Ayuda1.conexion VGCNx
  Me.Height = 3375
  Me.Width = 5505

End Sub

Private Sub axBotones_Click(Index As Integer)
    Select Case Index
        Case 0:
                Call Impresion
        
        Case 1: Unload Me
    
    End Select

End Sub

Sub Impresion()
'FIXIT: Declare 'arrform' and 'arrparm' con un tipo de datos de enlace en tiempo de compilación     FixIT90210ae-R1672-R1B8ZE
 Dim arrform(2) As Variant, arrparm() As Variant
    ReDim arrparm(2)
    arrparm(0) = VGParamSistem.BDEmpresa
    arrparm(1) = IIf(Ctr_Ayuda1.xclave = Empty, "%%", Trim$(Ctr_Ayuda1.xclave))
    Set VGvardllgen = New dllgeneral.dll_general
    arrform(0) = "@TituloReporte='" & "Listado de Cuentas Distribución" & "'"
    arrform(1) = "@Mes='" & VGvardllgen.DesMes(VGParamSistem.Mesproceso) & "'"
    Call ImpresionRptProc("rptListadoCtasDistribucion.rpt", arrform, arrparm)
End Sub
