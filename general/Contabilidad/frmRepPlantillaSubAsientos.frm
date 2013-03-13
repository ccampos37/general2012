VERSION 5.00
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Begin VB.Form frmRepPlantillaSubAsientos 
   Caption         =   "Reporte de Plantillas de SubAsientos"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5790
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   5790
   Begin VB.Frame Frame1 
      Caption         =   "Seleccionar"
      Height          =   1545
      Left            =   0
      TabIndex        =   2
      Top             =   330
      Width           =   5715
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuda1 
         Height          =   465
         Left            =   225
         TabIndex        =   3
         Top             =   390
         Width           =   5235
         _ExtentX        =   9234
         _ExtentY        =   820
         XcodMaxLongitud =   0
         xcodwith        =   800
         NomTabla        =   "ct_asiento"
         ListaCampos     =   "asientocodigo(1),asientodescripcion(1)"
         XcodCampo       =   "asientocodigo"
         XListCampo      =   "asientodescripcion"
         ListaCamposDescrip=   "Código,Descripción"
         ListaCamposText =   "asientocodigo,asientodescripcion"
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuda2 
         Height          =   465
         Left            =   225
         TabIndex        =   4
         Top             =   915
         Width           =   5220
         _ExtentX        =   9208
         _ExtentY        =   820
         XcodMaxLongitud =   0
         xcodwith        =   800
         NomTabla        =   "ct_subasiento"
         ListaCampos     =   "subasientocodigo(1),subasientodescripcion(1)"
         XcodCampo       =   "subasientocodigo"
         XListCampo      =   "subasientodescripcion"
         ListaCamposDescrip=   "Código,Descripción"
         ListaCamposText =   "subasientocodigo,subasientodescripcion"
      End
   End
   Begin VB.CommandButton cmdBotones 
      Caption         =   "&Cancelar"
      Height          =   360
      Index           =   1
      Left            =   2993
      TabIndex        =   1
      Top             =   2475
      Width           =   1215
   End
   Begin VB.CommandButton cmdBotones 
      Caption         =   "&Aceptar"
      Height          =   360
      Index           =   0
      Left            =   1583
      TabIndex        =   0
      Top             =   2460
      Width           =   1215
   End
End
Attribute VB_Name = "frmRepPlantillaSubAsientos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
  Ctr_Ayuda1.conexion VGCNx
  Ctr_Ayuda2.conexion VGCNx
  Me.Height = 3600
  Me.Width = 5910
End Sub

Private Sub cmdBotones_Click(Index As Integer)
  Select Case Index
    Case 0:
    
'FIXIT: Declare 'arrform' con un tipo de datos de enlace en tiempo de compilación          FixIT90210ae-R1672-R1B8ZE
       Dim arrform(2) As Variant, arrparm(5) As Variant
        arrparm(0) = VGParamSistem.BDEmpresa
        arrparm(1) = VGParametros.empresacodigo
        arrparm(2) = VGParamSistem.Anoproceso
        arrparm(3) = IIf(Ctr_Ayuda1.xclave = Empty, "%%", Ctr_Ayuda1.xclave)
        arrparm(4) = IIf(Ctr_Ayuda2.xclave = Empty, "%%", Ctr_Ayuda2.xclave)
        Set VGvardllgen = New dllgeneral.dll_general
        If Ctr_Ayuda1.xclave = Empty Then
            arrform(0) = "@TituloReporte='Reporte de Plantillas - Todos los Asientos'"
        Else
            arrform(0) = "@TituloReporte='" & "Reporte de Plantillas - Asiento: " & Ctr_Ayuda1.xclave & " " & Ctr_Ayuda1.xnombre & "'"
        End If
        arrform(1) = "@Mes='" & VGvardllgen.DesMes(VGParamSistem.Mesproceso) & "'"
        Call ImpresionRptProc("rptPlantillaSubAsiento.rpt", arrform, arrparm)
    
    Case 1: Unload Me
  
  
  End Select

End Sub

Private Sub Ctr_Ayuda1_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
  Ctr_Ayuda2.Filtro = "asientocodigo='" & Ctr_Ayuda1.xclave & "'"
End Sub
