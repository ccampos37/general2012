VERSION 5.00
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Begin VB.Form FrmSaldostotales 
   Caption         =   "Saldos Totales"
   ClientHeight    =   2595
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6510
   LinkTopic       =   "Form1"
   ScaleHeight     =   2595
   ScaleWidth      =   6510
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdCon 
      Caption         =   "Consultar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1560
      Width           =   1245
   End
   Begin VB.CommandButton CmdSal 
      Caption         =   "Retornar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1560
      Width           =   1245
   End
   Begin ctrlayuda_f.Ctr_Ayuda Ctr_ayulineas 
      Height          =   315
      Left            =   960
      TabIndex        =   4
      Top             =   960
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   556
      XcodMaxLongitud =   0
      xcodwith        =   500
      NomTabla        =   "lineas"
      ListaCampos     =   "lin_codigo(1),lin_nombre(1)"
      XcodCampo       =   "lin_codigo"
      XListCampo      =   "lin_nombre"
      ListaCamposDescrip=   "Código,Descripción"
      ListaCamposText =   "lin_codigo,lin_nombre"
      Requerido       =   0   'False
   End
   Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyuFamilia 
      Height          =   315
      Left            =   960
      TabIndex        =   2
      Top             =   570
      Width           =   5250
      _ExtentX        =   9260
      _ExtentY        =   556
      XcodMaxLongitud =   0
      xcodwith        =   500
      NomTabla        =   "familia"
      ListaCampos     =   "fam_codigo(1),fam_nombre(1)"
      XcodCampo       =   "fam_codigo"
      XListCampo      =   "fam_nombre"
      ListaCamposDescrip=   "Código,Descripción"
      ListaCamposText =   "fam_codigo,fam_nombre"
      Requerido       =   0   'False
   End
   Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayugrupo 
      Height          =   315
      Left            =   945
      TabIndex        =   7
      Top             =   1320
      Visible         =   0   'False
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   556
      XcodMaxLongitud =   0
      xcodwith        =   500
      NomTabla        =   "lineas"
      ListaCampos     =   "lin_codigo(1),lin_nombre(1)"
      XcodCampo       =   "lin_codigo"
      XListCampo      =   "lin_nombre"
      ListaCamposDescrip=   "Código,Descripción"
      ListaCamposText =   "lin_codigo,lin_nombre"
      Requerido       =   0   'False
   End
   Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyuTipo 
      Height          =   315
      Left            =   960
      TabIndex        =   1
      Top             =   120
      Width           =   5250
      _ExtentX        =   9260
      _ExtentY        =   556
      XcodMaxLongitud =   0
      xcodwith        =   500
      NomTabla        =   "tipo_articulo"
      ListaCampos     =   "COD_TIPO(1),DES_TIPO(1)"
      XcodCampo       =   "COD_TIPO"
      XListCampo      =   "DES_TIPO"
      ListaCamposDescrip=   "Código,Descripción"
      ListaCamposText =   "COD_TIPO,DES_TIPO"
      Requerido       =   0   'False
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Tipo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   150
      Width           =   360
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Grupo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   1440
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Familia"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   600
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "LInea"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   135
      TabIndex        =   5
      Top             =   1065
      Width           =   480
   End
End
Attribute VB_Name = "FrmSaldostotales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset

Private Sub CmdCon_Click()
Dim aparam(4) As Variant
Dim aform(2) As Variant

aform(0) = "fecha='" & Date & "'"
aform(1) = "tipo='" & Ctr_AyuTipo.xnombre & "'"
aparam(0) = VGCNx.DefaultDatabase
aparam(1) = IIf(Ctr_AyuTipo.xclave = "", "%%", Ctr_AyuTipo.xclave)
If Ctr_AyuTipo.xclave <> "" Then
  aparam(2) = IIf(Ctr_AyuFamilia.xclave = "", " ", Ctr_AyuFamilia.xclave)
  aparam(3) = IIf(Ctr_ayulineas.xclave = "", " ", Ctr_ayulineas.xclave)
Else
  aparam(2) = IIf(Ctr_AyuFamilia.xclave = "", "%%", Ctr_AyuFamilia.xclave)
  aparam(3) = IIf(Ctr_ayulineas.xclave = "", "%%", Ctr_ayulineas.xclave)
End If

Call ImpresionRptProc("al_saldosconsolidadosxTipo.rpt", aform, aparam, , "Saldos Consolidados x tipo de Articulo")

End Sub

Private Sub CmdSal_Click()
Unload Me
End Sub


Private Sub Ctr_AyuFamilia_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
Set rs = Nothing
Set rs = VGCNx.Execute(" select * from lineas where fam_codigo='" & Ctr_AyuFamilia.xclave & "'")
If rs.RecordCount > 0 Then
   Ctr_ayulineas.filtro = "fam_codigo='" & Ctr_AyuFamilia.xclave & "'"
   Ctr_ayulineas.Visible = True
 Else
  Ctr_ayulineas.xclave = ""
  Ctr_ayulineas.Visible = True
End If
End Sub



Private Sub Form_Load()
Call Ctr_AyuFamilia.Conexion(VGCNx)
Call Ctr_ayulineas.Conexion(VGCNx)
Call Ctr_Ayugrupo.Conexion(VGCNx)
Call Ctr_AyuTipo.Conexion(VGCNx)

Ctr_ayulineas.Visible = False
Ctr_Ayugrupo.Visible = False

CmdSal.Picture = MDIPrincipal.ImageList2.ListImages("Retornar").Picture
CmdCon.Picture = MDIPrincipal.ImageList2.ListImages("Facturado").Picture

End Sub




