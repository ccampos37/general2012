VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Begin VB.Form FrmRepTraslados 
   Caption         =   "Informe de Traslados"
   ClientHeight    =   3330
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6645
   LinkTopic       =   "Form2"
   ScaleHeight     =   3330
   ScaleWidth      =   6645
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      Caption         =   "Transaccion"
      Height          =   975
      Left            =   240
      TabIndex        =   12
      Top             =   2040
      Width           =   3015
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyuTransaccion 
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   2760
         _ExtentX        =   4868
         _ExtentY        =   661
         XcodMaxLongitud =   0
         xcodwith        =   100
         NomTabla        =   "tabtransa"
         TituloAyuda     =   "Transacciones"
         ListaCampos     =   "TT_CODMOV(1), TT_DESCRI(1)"
         XcodCampo       =   "tt_codmov"
         XListCampo      =   "tt_DESCRI"
         ListaCamposDescrip=   "Codigo,Descripcion"
         ListaCamposText =   "TT_CODMOV, TT_DESCRI"
         Requerido       =   0   'False
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Almacen Destino"
      Height          =   735
      Left            =   3360
      TabIndex        =   10
      Top             =   1200
      Width           =   3015
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyuAlmacen2 
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   2760
         _ExtentX        =   4868
         _ExtentY        =   661
         XcodMaxLongitud =   0
         xcodwith        =   100
         NomTabla        =   "tabalm"
         TituloAyuda     =   "Almacenes"
         ListaCampos     =   "TAALMA(1),TADESCRI(1)"
         XcodCampo       =   "TAALMA"
         XListCampo      =   "TADESCRI"
         ListaCamposDescrip=   "Codigo,Descripcion"
         ListaCamposText =   "TAALMA,TADESCRI"
         Requerido       =   0   'False
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Almacen Origen"
      Height          =   735
      Left            =   240
      TabIndex        =   8
      Top             =   1200
      Width           =   3015
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyuAlmacen1 
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   2760
         _ExtentX        =   4868
         _ExtentY        =   661
         XcodMaxLongitud =   0
         xcodwith        =   100
         NomTabla        =   "tabalm"
         TituloAyuda     =   "Almacenes"
         ListaCampos     =   "TAALMA(1),TADESCRI(1)"
         XcodCampo       =   "TAALMA"
         XListCampo      =   "TADESCRI"
         ListaCamposDescrip=   "Codigo,Descripcion"
         ListaCamposText =   "TAALMA,TADESCRI"
         Requerido       =   0   'False
      End
   End
   Begin VB.Frame Frame3 
      Height          =   975
      Left            =   3360
      TabIndex        =   5
      Top             =   2040
      Width           =   3015
      Begin VB.CommandButton Command1 
         Caption         =   "&Aceptar"
         Height          =   645
         Left            =   480
         Picture         =   "FrmRepTraslados.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   795
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Cancelar"
         Height          =   645
         Left            =   1920
         Picture         =   "FrmRepTraslados.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   795
      End
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   150
      Width           =   6135
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   1080
         TabIndex        =   1
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   99221505
         CurrentDate     =   36437
         MaxDate         =   401768
         MinDate         =   35796
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   4080
         TabIndex        =   2
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   99221505
         CurrentDate     =   36437
         MaxDate         =   401768
         MinDate         =   35796
      End
      Begin VB.Label Label6 
         Caption         =   "Hasta"
         Height          =   255
         Left            =   3240
         TabIndex        =   4
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Desde"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   975
      End
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   0
      Top             =   1470
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "FrmRepTraslados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Dim aparam(6) As Variant
Dim aform(3) As Variant
Dim titulo As String
Dim VGDllGeneral As New dllgeneral.dll_general
If Ctr_AyuAlmacen1.xclave = "" Then
   MsgBox (" Debe ingresar Codigo de Almacen ")
   Ctr_AyuAlmacen1.SetFocus
   Exit Sub
End If
titulo = "al_transferencias  -- Documentos Emitidos Detallados"
aparam(0) = Trim(VGCNx.DefaultDatabase)
aparam(1) = Format(DTPicker1.Value, "dd/mm/yyyy")
aparam(2) = Format(DTPicker2.Value, "dd/mm/yyyy")
aparam(3) = Ctr_AyuAlmacen1.xclave
If Ctr_AyuAlmacen2.xclave = "" Then
   aparam(4) = "%%"
 Else
   aparam(4) = Ctr_AyuAlmacen2.xclave
End If
If Ctr_AyuTransaccion.xclave = "" Then
   aparam(5) = "%%"
 Else
   aparam(5) = Ctr_AyuTransaccion.xclave
End If
aform(0) = "fechainicio ='" & DTPicker1 & "'"
aform(1) = "fechafin ='" & DTPicker2 & "'"
aform(2) = "fechafin ='" & DTPicker2 & "'"

Call ImpresionRptProc("al_transferencias.rpt", aform, aparam, , titulo)
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
  DTPicker1.Value = Date
  DTPicker2.Value = Date
  Call Ctr_AyuAlmacen1.Conexion(VGCNx)
  Call Ctr_AyuAlmacen2.Conexion(VGCNx)
  Call Ctr_AyuTransaccion.Conexion(VGCNx): Ctr_AyuTransaccion.filtro = " tt_codtrans_auto<>''"
End Sub


