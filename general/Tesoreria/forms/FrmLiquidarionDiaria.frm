VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Begin VB.Form FrmLiquidarionDiaria 
   Caption         =   "Form1"
   ClientHeight    =   4470
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6990
   LinkTopic       =   "Form1"
   ScaleHeight     =   4470
   ScaleWidth      =   6990
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Tipo"
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   240
      TabIndex        =   6
      Top             =   45
      Width           =   6495
      Begin VB.OptionButton OptResumido 
         Caption         =   "Resumido"
         Height          =   255
         Left            =   3960
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton OptDetallado 
         Caption         =   "Detallado"
         Height          =   255
         Left            =   1080
         TabIndex        =   7
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame6 
      Height          =   1215
      Left            =   960
      TabIndex        =   3
      Top             =   2925
      Width           =   4695
      Begin VB.CommandButton cmdaceptar 
         Caption         =   "&Aceptar"
         Height          =   615
         Left            =   480
         TabIndex        =   5
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton cmdsalir 
         Caption         =   "&Cancelar"
         Height          =   615
         Left            =   2640
         TabIndex        =   4
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Filtrar Por"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1815
      Left            =   240
      TabIndex        =   0
      Top             =   765
      Width           =   6525
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuempresa 
         Height          =   375
         Left            =   1320
         TabIndex        =   1
         Top             =   360
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   661
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
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyudaCaja 
         Height          =   375
         Left            =   1320
         TabIndex        =   9
         Top             =   840
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   661
         XcodMaxLongitud =   11
         xcodwith        =   400
         NomTabla        =   "te_codigocaja"
         TituloAyuda     =   "Busqueda de Caja"
         ListaCampos     =   "cajacodigo(1),cajadescripcion(1)"
         XcodCampo       =   "cajacodigo"
         XListCampo      =   "cajadescripcion"
         ListaCamposDescrip=   "Codigo,Descripcion"
         ListaCamposText =   "cajacodigo,cajadescripcion"
      End
      Begin MSComCtl2.DTPicker DTPickerFecInicio 
         Height          =   375
         Left            =   1320
         TabIndex        =   13
         Top             =   1320
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   98500609
         CurrentDate     =   41335
      End
      Begin MSComCtl2.DTPicker DTPickerFecFinal 
         Height          =   375
         Left            =   4320
         TabIndex        =   14
         Top             =   1320
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   98500609
         CurrentDate     =   41335
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta :"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   3585
         TabIndex        =   12
         Top             =   1395
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Cod. Caja"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   11
         Top             =   840
         Width           =   885
      End
      Begin VB.Label lbMon 
         Caption         =   "Desde :"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   465
         TabIndex        =   10
         Top             =   1365
         Width           =   735
      End
      Begin VB.Label Lblempresa 
         AutoSize        =   -1  'True
         Caption         =   "Empresa :"
         Height          =   195
         Left            =   345
         TabIndex        =   2
         Top             =   450
         Width           =   705
      End
   End
End
Attribute VB_Name = "FrmLiquidarionDiaria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sresumido As Integer

Private Sub cmdaceptar_Click()
Dim arrform(12) As Variant
Dim arrparm(6) As Variant
    
arrparm(0) = VGParamSistem.BDEmpresa
arrparm(1) = Ctr_Ayuempresa.xclave
arrparm(2) = IIf(Ctr_AyudaCaja.xclave = Empty, "%%", Trim(Ctr_AyudaCaja.xclave))
arrparm(3) = Format(DTPickerFecInicio.Value, "dd/mm/yyyy")
arrparm(4) = Format(DTPickerFecFinal.Value, "dd/mm/yyyy")

arrform(0) = "rangofecha=' DEL : " & Format(DTPickerFecInicio.Value, "dd/mm/yyyy") & " AL " & Format(DTPickerFecFinal.Value, "dd/mm/yyyy") & "'"
If Optdetallado Then
    arrform(1) = "tipo=' DETALLADO '"
    Call ImpresionRptProc("te_liquidacionDiaria.rpt", arrform, arrparm, , "Detallado")
   Else
    arrform(1) = "tipo=' RESUMIDO '"
    Call ImpresionRptProc("te_liquidacionDiariaresumen.rpt", arrform, arrparm, , "resumido")
End If
End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
  Dim cFecha As Date
  DTPickerFecInicio.Value = VGParamSistem.fechatrabajo
  DTPickerFecFinal.Value = VGParamSistem.fechatrabajo
  Call Ctr_Ayuempresa.Conexion(VGCNx)
  If VGParametros.sistemamultiempresas = False Then
     Ctr_Ayuempresa.xclave = VGParametros.empresacodigo
     Ctr_Ayuempresa.Ejecutar
     Ctr_Ayuempresa.Enabled = False
  End If
  Call Ctr_AyudaCaja.Conexion(VGCNx)
  If sresumido = 1 Then
     Optresumido = True
     Ctr_AyudaCaja.Enabled = False
   Else
     Optdetallado = True
     If VGParametros.listacajas <> "" Then
        SQL = " CAJACODIGO IN (" & VGParametros.listacajas & ")"
        Ctr_AyudaCaja.filtro = SQL
        Ctr_AyudaCaja.Enabled = True
     End If
  End If
End Sub

Public Property Let resumido(ByVal vNewValue As Variant)
sresumido = vNewValue
End Property
