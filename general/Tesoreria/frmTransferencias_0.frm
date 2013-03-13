VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Object = "{4D137D9C-00A6-4458-9B46-7E95DB76D55B}#9.0#0"; "TextFer.ocx"
Begin VB.Form frmTransferencias 
   Caption         =   "Transferencias"
   ClientHeight    =   4650
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12930
   LinkTopic       =   "Form1"
   ScaleHeight     =   4650
   ScaleWidth      =   12930
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame5 
      Caption         =   "DATOS REFERENCIA X RENDIR"
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
      Height          =   1575
      Left            =   5760
      TabIndex        =   40
      Top             =   3000
      Width           =   5535
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ay_Transf 
         Height          =   330
         Left            =   1395
         TabIndex        =   14
         Top             =   465
         Width           =   3390
         _ExtentX        =   5980
         _ExtentY        =   582
         XcodMaxLongitud =   0
         xcodwith        =   500
         NomTabla        =   "te_conceptocaja"
         ListaCampos     =   "conceptocodigo(1),conceptodescripcion(1),conceptoingresaref(1)"
         XcodCampo       =   "conceptocodigo"
         XListCampo      =   "conceptodescripcion"
         ListaCamposDescrip=   "Código,Descripción,Entidad"
         ListaCamposText =   "conceptocodigo,conceptodescripcion,conceptoingresaref"
         Requerido       =   0   'False
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyuEntidad 
         Height          =   330
         Left            =   1395
         TabIndex        =   15
         Top             =   1080
         Width           =   3990
         _ExtentX        =   7038
         _ExtentY        =   582
         XcodMaxLongitud =   0
         xcodwith        =   1000
         NomTabla        =   "ct_entidad"
         ListaCampos     =   "entidadcodigo(1),entidadrazonsocial(1)"
         XcodCampo       =   "entidadcodigo"
         XListCampo      =   "entidadrazonsocial"
         ListaCamposDescrip=   "Código,Descripción"
         ListaCamposText =   "entidadcodigo,entidadrazonsocial"
         Requerido       =   0   'False
      End
      Begin VB.Label Label1 
         Caption         =   "Responsable"
         Height          =   270
         Index           =   5
         Left            =   120
         TabIndex        =   42
         Top             =   1110
         Width           =   1365
      End
      Begin VB.Label Label1 
         Caption         =   "Cpto Transferencia"
         Height          =   390
         Index           =   10
         Left            =   120
         TabIndex        =   41
         Top             =   510
         Width           =   1320
      End
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Index           =   2
      Left            =   10680
      TabIndex        =   38
      Top             =   45
      Width           =   1200
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   6480
      TabIndex        =   37
      Top             =   45
      Width           =   1200
   End
   Begin VB.Frame Frame6 
      Height          =   1575
      Left            =   11400
      TabIndex        =   32
      Top             =   3000
      Width           =   1335
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   390
         Left            =   120
         TabIndex        =   34
         Top             =   375
         Width           =   1080
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   390
         Left            =   120
         TabIndex        =   33
         Top             =   990
         Width           =   1080
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "DATOS"
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
      Height          =   2355
      Left            =   5760
      TabIndex        =   27
      Top             =   360
      Width           =   6975
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   315
         Left            =   1770
         TabIndex        =   11
         Top             =   1110
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   556
         _Version        =   393216
         Format          =   16777217
         CurrentDate     =   37628
      End
      Begin TextFer.TxFer txt 
         Height          =   300
         Index           =   1
         Left            =   1770
         TabIndex        =   9
         Top             =   240
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   529
         Object.CausesValidation=   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   ""
         ColorIlumina    =   -2147483624
         SaltarAlEnter   =   -1  'True
         Valor           =   ""
         TipoDato        =   1
         NumeroDecimales =   4
      End
      Begin TextFer.TxFer txt 
         Height          =   300
         Index           =   2
         Left            =   1770
         TabIndex        =   10
         Top             =   615
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   529
         Object.CausesValidation=   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   ""
         ColorIlumina    =   -2147483624
         SaltarAlEnter   =   -1  'True
         Valor           =   ""
         TipoDato        =   1
         SignodeMiles    =   -1  'True
         NumeroDecimales =   4
         MarcarTextoAlEnfoque=   -1  'True
      End
      Begin TextFer.TxFer txt 
         Height          =   300
         Index           =   3
         Left            =   1770
         TabIndex        =   13
         Top             =   1980
         Width           =   4650
         _ExtentX        =   8202
         _ExtentY        =   529
         Object.CausesValidation=   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   50
         Text            =   ""
         ColorIlumina    =   -2147483624
         SaltarAlEnter   =   -1  'True
         Valor           =   ""
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayutransf 
         Height          =   315
         Left            =   1755
         TabIndex        =   12
         Top             =   1575
         Visible         =   0   'False
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   556
         XcodMaxLongitud =   7
         xcodwith        =   800
         NomTabla        =   "te_cabecerarecibos"
         TituloAyuda     =   "Busqueda de Documentos x rendir"
         ListaCampos     =   "cabrec_numreciboegreso(1),cabrec_descripcion(1),SaldoDocxRendir(1)"
         XcodCampo       =   "cabrec_numreciboegreso"
         XListCampo      =   "cabrec_descripcion"
         ListaCamposDescrip=   "Nro.transferencia,descripcion,Saldo"
         ListaCamposText =   "cabrec_numreciboegreso,cabrec_descripcion,SaldoDocxRendir"
      End
      Begin VB.Label LeReferencia 
         AutoSize        =   -1  'True
         Caption         =   "Nro.Transf."
         Height          =   195
         Left            =   120
         TabIndex        =   43
         Top             =   1575
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo de Cambio"
         Height          =   270
         Index           =   6
         Left            =   120
         TabIndex        =   31
         Top             =   300
         Width           =   1125
      End
      Begin VB.Label Label1 
         Caption         =   "Importe a Transferir"
         Height          =   300
         Index           =   7
         Left            =   150
         TabIndex        =   30
         Top             =   660
         Width           =   1410
      End
      Begin VB.Label Label1 
         Caption         =   "Observaciones"
         Height          =   270
         Index           =   8
         Left            =   120
         TabIndex        =   29
         Top             =   2010
         Width           =   2085
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Transferencia"
         Height          =   390
         Index           =   9
         Left            =   120
         TabIndex        =   28
         Top             =   1050
         Width           =   1365
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "DESTINO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1575
      Left            =   240
      TabIndex        =   23
      Top             =   3015
      Width           =   5415
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ay_Destino 
         Height          =   330
         Left            =   1125
         TabIndex        =   7
         Top             =   600
         Width           =   3360
         _ExtentX        =   5927
         _ExtentY        =   582
         XcodMaxLongitud =   0
         xcodwith        =   500
         Requerido       =   0   'False
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ay_CtaMonedaDestino 
         Height          =   330
         Left            =   1125
         TabIndex        =   8
         Top             =   1095
         Width           =   3120
         _ExtentX        =   5503
         _ExtentY        =   582
         XcodMaxLongitud =   0
         xcodwith        =   500
         Requerido       =   0   'False
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuempresadestino 
         Height          =   315
         Left            =   1125
         TabIndex        =   6
         Top             =   180
         Width           =   4140
         _ExtentX        =   7303
         _ExtentY        =   556
         XcodMaxLongitud =   3
         xcodwith        =   300
         NomTabla        =   "co_multiempresas"
         TituloAyuda     =   "Busqueda de Empresas"
         ListaCampos     =   "empresacodigo(1),empresadescripcion(1)"
         XcodCampo       =   "empresacodigo"
         XListCampo      =   "empresadescripcion"
         ListaCamposDescrip=   "Codigo,Descripcion"
         ListaCamposText =   "empresacodigo,empresadescripcion"
      End
      Begin VB.Label Lblempresa2 
         AutoSize        =   -1  'True
         Caption         =   "Empresa :"
         Height          =   195
         Left            =   135
         TabIndex        =   44
         Top             =   240
         Width           =   705
      End
      Begin VB.Label lblMonDes 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   4410
         TabIndex        =   26
         Top             =   1125
         Width           =   585
      End
      Begin VB.Label Label1 
         Caption         =   "Destino"
         Height          =   270
         Index           =   3
         Left            =   120
         TabIndex        =   25
         Top             =   630
         Width           =   2085
      End
      Begin VB.Label Label1 
         Caption         =   "Cuenta  / Moneda"
         Height          =   360
         Index           =   4
         Left            =   120
         TabIndex        =   24
         Top             =   1035
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   15
      Left            =   1200
      TabIndex        =   22
      Top             =   6840
      Width           =   135
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   1650
      TabIndex        =   19
      Top             =   45
      Width           =   1680
   End
   Begin VB.Frame Frame1 
      Caption         =   "ORIGEN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2355
      Left            =   180
      TabIndex        =   4
      Top             =   360
      Width           =   5475
      Begin TextFer.TxFer txt 
         Height          =   300
         Index           =   0
         Left            =   1305
         TabIndex        =   5
         Top             =   1890
         Width           =   3945
         _ExtentX        =   6959
         _ExtentY        =   529
         Object.CausesValidation=   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   16
         Text            =   ""
         ColorIlumina    =   -2147483624
         SaltarAlEnter   =   -1  'True
         Valor           =   ""
         ColorTextoAlEnfocar=   -2147483640
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ay_Origen 
         Height          =   330
         Left            =   1305
         TabIndex        =   1
         Top             =   615
         Width           =   4035
         _ExtentX        =   7117
         _ExtentY        =   582
         XcodMaxLongitud =   0
         xcodwith        =   500
         NomTabla        =   "te_codigocaja"
         ListaCampos     =   "cajacodigo(1),cajadescripcion(1)"
         XcodCampo       =   "cajacodigo"
         XListCampo      =   "cajadescripcion"
         ListaCamposDescrip=   "Código,Descripción"
         ListaCamposText =   "cajacodigo,cajadescripcion "
         Requerido       =   0   'False
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ay_CtaMonedaOrigen 
         Height          =   360
         Left            =   1305
         TabIndex        =   2
         Top             =   1080
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   635
         XcodMaxLongitud =   0
         xcodwith        =   500
         Requerido       =   0   'False
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyuDocumento 
         Height          =   315
         Left            =   1305
         TabIndex        =   3
         Top             =   1455
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   556
         XcodMaxLongitud =   2
         xcodwith        =   200
         NomTabla        =   "cp_tipodocumento"
         TituloAyuda     =   "Busqueda de Tipo de  Documento"
         ListaCampos     =   "tdocumentocodigo(1),tdocumentodescripcion(1),tdocumentotipo(1),documentoretencion(1)"
         XcodCampo       =   "tdocumentocodigo"
         XListCampo      =   "tdocumentodescripcion"
         ListaCamposDescrip=   "Código,Descripción,CargoAbono,Retencion"
         ListaCamposText =   "tdocumentocodigo,tdocumentodescripcion,tdocumentotipo,documentoretencion"
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuempresaorigen 
         Height          =   315
         Left            =   1305
         TabIndex        =   0
         Top             =   225
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   556
         XcodMaxLongitud =   3
         xcodwith        =   300
         NomTabla        =   "co_multiempresas"
         TituloAyuda     =   "Busqueda de Empresas"
         ListaCampos     =   "empresacodigo(1),empresadescripcion(1)"
         XcodCampo       =   "empresacodigo"
         XListCampo      =   "empresadescripcion"
         ListaCamposDescrip=   "Codigo,Descripcion"
         ListaCamposText =   "empresacodigo,empresadescripcion"
      End
      Begin VB.Label Lblempresa1 
         AutoSize        =   -1  'True
         Caption         =   "Empresa :"
         Height          =   195
         Left            =   180
         TabIndex        =   45
         Top             =   285
         Width           =   705
      End
      Begin VB.Label Label1 
         Caption         =   "Doc. de Transf."
         Height          =   405
         Index           =   11
         Left            =   120
         TabIndex        =   39
         Top             =   1485
         Width           =   1020
      End
      Begin VB.Label lblMonOrigen 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   4680
         TabIndex        =   20
         Top             =   1080
         Width           =   600
      End
      Begin VB.Label Label1 
         Caption         =   "Cheque/ Referencia"
         Height          =   360
         Index           =   2
         Left            =   105
         TabIndex        =   18
         Top             =   1875
         Width           =   1440
      End
      Begin VB.Label Label1 
         Caption         =   "Cuenta o Moneda"
         Height          =   420
         Index           =   1
         Left            =   150
         TabIndex        =   17
         Top             =   1020
         Width           =   960
      End
      Begin VB.Label Label1 
         Caption         =   "Origen"
         Height          =   255
         Index           =   0
         Left            =   150
         TabIndex        =   16
         Top             =   675
         Width           =   870
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Nro Recibo Egreso"
      Height          =   180
      Index           =   2
      Left            =   4800
      TabIndex        =   36
      Top             =   75
      Width           =   1440
   End
   Begin VB.Label Label2 
      Caption         =   "Nro Recibo Ingreso"
      Height          =   180
      Index           =   1
      Left            =   8880
      TabIndex        =   35
      Top             =   75
      Width           =   1440
   End
   Begin VB.Label Label2 
      Caption         =   "Nro Transferencia"
      Height          =   180
      Index           =   0
      Left            =   270
      TabIndex        =   21
      Top             =   75
      Width           =   1440
   End
End
Attribute VB_Name = "frmTransferencias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim m_CasoOrigen As String
Dim m_CasoDestino As String
Dim m_cuentasxrendir As Integer
Dim m_fondofijo As Integer
Dim m_proveedor As String
Dim m_tipo As Integer
Dim entidad As Integer
Dim cambio As Integer
Property Let titulo(valor As String)
frmTransferencias.Caption = valor
End Property

Property Let CasoOrigen(valor As String)
   m_CasoOrigen = valor
End Property
Property Let tipo(valor As String)
   m_tipo = valor
End Property
Property Let CasoDestino(valor As String)
   m_CasoDestino = valor
End Property
Property Let cuentasxrendir(valor As String)
   m_cuentasxrendir = valor
End Property
Property Let fondofijo(valor As String)
   m_fondofijo = valor
End Property
Property Let Proveedor(valor As String)
   m_proveedor = valor
End Property


Private Sub Ctr_Ay_Transf_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
entidad = ColecCampos("conceptoingresaref")
If entidad = 1 Then
   Ctr_AyuEntidad.Visible = True
 Else
   Ctr_AyuEntidad.Visible = False
End If
End Sub

Private Sub Form_Load()
   
   Call Ctr_AyuDocumento.Conexion(VGCNx): Ctr_AyuDocumento.Filtro = "tdocumentotipo='A' and tdocumentovalidabanco=1 and  tdocumentoingplan=0"
   Call Ctr_AyuEntidad.Conexion(VGCNx)
   Call Ctr_Ayutransf.Conexion(VGCNx)
   Call Ctr_Ayuempresaorigen.Conexion(VGCNx)
   Call Ctr_Ayuempresadestino.Conexion(VGCNx)
   
   Call MuestraNumeradorTransf
   Call ActivarCajaBancoOrigen
   Call ActivarCajaBancoDestino
   Call ActivarCtaMonedaOrigen
   Call ActivarCtaMonedaDestino
   Call ActivarConceptoTransf
   
   DTPicker1.Value = Format(Now, "dd/mm/yyyy")
   If (m_fondofijo = 1 Or m_cuentasxrendir = 1) And m_tipo = 1 Then
      Ctr_Ayutransf.Visible = True
      LeReferencia.Visible = True
      cambio = 1
    Else
      Ctr_Ayutransf.Visible = False
      LeReferencia.Visible = False
      cambio = 0
   End If
   If VGParametros.sistemamultiempresas Then
      Ctr_Ayuempresaorigen.Enabled = True
      Ctr_Ayuempresadestino.Enabled = True
    Else
      Ctr_Ayuempresaorigen.xclave = VGParametros.empresacodigo: Ctr_Ayuempresaorigen.Ejecutar
      Ctr_Ayuempresaorigen.Enabled = False
      Ctr_Ayuempresadestino.xclave = VGParametros.empresacodigo: Ctr_Ayuempresadestino.Ejecutar
      Ctr_Ayuempresadestino.Enabled = False
   End If
End Sub

Private Sub cmdaceptar_Click()
On Error GoTo X
   
   If ValidarData = True Then
     VGCNx.BeginTrans
     Call ActualizaNumerador
     Call GrabarDataOrigen
     Call NumeradorIngreso
     Call GrabarDataDestino
 '    Call GeneraAsientoEnlineaTesorTransfer(DTPicker1.Value, Trim(Text1(1).Text))
      
     
     VGCNx.CommitTrans
     Call LimpiarForm
     Call ImpresionTransferencias(Text1(1).Text)
     Call MuestraNumeradorTransf
     Ctr_Ay_Origen.SetFocus
   End If
   Exit Sub
   
X:
    MsgBox "La Grabación de la Transferencia no se pudo Completar" & Chr(13) & "Error: " & Err.Number & " - " & Err.Description, vbInformation, Caption
    VGCNx.RollbackTrans
    Exit Sub
    Resume
   
End Sub

Private Sub Ctr_Ay_CtaMonedaOrigen_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
  Select Case m_CasoOrigen
    Case "C":
      lblMonOrigen.Caption = Ctr_Ay_CtaMonedaOrigen.xclave
    Case "B":
      lblMonOrigen.Caption = ColecCampos("monedacodigo").Value
      txt(0).Text = ColecCampos("cbanco_nrocheque").Value
  End Select

End Sub

Private Sub Ctr_Ay_Destino_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
  Select Case m_CasoDestino
    Case "C":
        
    Case "B":
       Ctr_Ay_CtaMonedaDestino.Filtro = "cbanco_codigo='" & ColecCampos("bancocodigo").Value & "'"
  End Select
  
End Sub

Private Sub Ctr_Ay_Origen_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
  Select Case m_CasoOrigen
    Case "B":
      Ctr_Ay_CtaMonedaOrigen.Filtro = "cbanco_codigo='" & ColecCampos("bancocodigo").Value & "'"
  End Select

End Sub

Private Sub Ctr_Ay_CtaMonedaDestino_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
  Select Case m_CasoDestino
    Case "C":
      lblMonDes.Caption = Ctr_Ay_CtaMonedaDestino.xclave
    Case "B":
      lblMonDes.Caption = ColecCampos("monedacodigo").Value
  End Select

End Sub

Private Sub CmdCancelar_Click()
  Unload Me
End Sub

Sub LimpiarForm()
  Dim I As Integer
  For I = 0 To 3
    txt(I).Text = Empty
  Next I
  Ctr_Ay_Origen.xclave = Empty
  Ctr_Ay_Destino.xclave = Empty
  Ctr_Ay_CtaMonedaOrigen.xclave = Empty
  Ctr_Ay_CtaMonedaDestino.xclave = Empty
  Ctr_Ay_Transf.xclave = Empty
  lblMonOrigen.Caption = Empty
  lblMonDes.Caption = Empty
  Ctr_Ay_Origen.xnombre = Empty
  Ctr_Ay_Destino.xnombre = Empty
  Ctr_Ay_CtaMonedaOrigen.xnombre = Empty
  Ctr_Ay_CtaMonedaDestino.xnombre = Empty
  Ctr_Ay_Transf.xnombre = Empty
  Text1(0).Text = Empty
  'Text1(1).Text = Empty
End Sub

Sub ActivarCajaBancoOrigen()
  Select Case m_CasoOrigen
    Case "C":
      Ctr_Ay_Origen.ListaCampos = "cajacodigo(1),cajadescripcion(1)"
      Ctr_Ay_Origen.ListaCamposDescrip = "Código,Descripción"
      Ctr_Ay_Origen.ListaCamposText = "cajacodigo,cajadescripcion"
      Ctr_Ay_Origen.NomTabla = "te_codigocaja"
      Ctr_Ay_Origen.XcodCampo = "cajacodigo"
      Ctr_Ay_Origen.XListCampo = "cajadescripcion"
      Ctr_Ay_Origen.Conexion VGCNx
      If m_tipo = 0 Then
        Ctr_Ay_Origen.Filtro = " not (isnull(CajaCuentaxRendir,0)=1 or isnull(Cajafondofijo,0)=1 )"
       Else
        Ctr_Ay_Origen.Filtro = " isnull(CajaCuentaxRendir,0)=" & m_cuentasxrendir & " and isnull(Cajafondofijo,0)=" & m_fondofijo
      End If
    Case "B":
      Ctr_Ay_Origen.ListaCampos = "bancocodigo(1),bancodescripcion(1)"
      Ctr_Ay_Origen.ListaCamposDescrip = "Código,Descripción"
      Ctr_Ay_Origen.ListaCamposText = "bancocodigo,bancodescripcion"
      Ctr_Ay_Origen.NomTabla = "gr_banco"
      Ctr_Ay_Origen.XcodCampo = "bancocodigo"
      Ctr_Ay_Origen.XListCampo = "bancodescripcion"
      Ctr_Ay_Origen.Conexion VGCNx
  End Select
End Sub

Sub ActivarCajaBancoDestino()
  Select Case m_CasoDestino
    Case "C":
      Ctr_Ay_Destino.ListaCampos = "cajacodigo(1),cajadescripcion(1)"
      Ctr_Ay_Destino.ListaCamposDescrip = "Código,Descripción"
      Ctr_Ay_Destino.ListaCamposText = "cajacodigo,cajadescripcion"
      Ctr_Ay_Destino.NomTabla = "te_codigocaja"
      Ctr_Ay_Destino.XcodCampo = "cajacodigo"
      Ctr_Ay_Destino.XListCampo = "cajadescripcion"
      Ctr_Ay_Destino.Conexion VGCNx
      If m_tipo = 1 Then
         Ctr_Ay_Destino.Filtro = " not (isnull(CajaCuentaxRendir,0)=1 or isnull(Cajafondofijo,0)=1 )"
       Else
         Ctr_Ay_Destino.Filtro = " isnull(CajaCuentaxRendir,0)=" & m_cuentasxrendir & " and isnull(Cajafondofijo,0)=" & m_fondofijo
      End If
    Case "B":
      Ctr_Ay_Destino.ListaCampos = "bancocodigo(1),bancodescripcion(1)"
      Ctr_Ay_Destino.ListaCamposDescrip = "Código,Descripción"
      Ctr_Ay_Destino.ListaCamposText = "bancocodigo,bancodescripcion"
      Ctr_Ay_Destino.NomTabla = "gr_banco"
      Ctr_Ay_Destino.XcodCampo = "bancocodigo"
      Ctr_Ay_Destino.XListCampo = "bancodescripcion"
      Ctr_Ay_Destino.Conexion VGCNx
  End Select
End Sub

Sub ActivarCtaMonedaOrigen()
  Select Case m_CasoOrigen
    Case "C":
      Ctr_Ay_CtaMonedaOrigen.ListaCampos = "monedacodigo(1),monedadescripcion(1)"
      Ctr_Ay_CtaMonedaOrigen.ListaCamposDescrip = "Código,Descripción"
      Ctr_Ay_CtaMonedaOrigen.ListaCamposText = "monedacodigo,monedadescripcion"
      Ctr_Ay_CtaMonedaOrigen.NomTabla = "gr_moneda"
      Ctr_Ay_CtaMonedaOrigen.XcodCampo = "monedacodigo"
      Ctr_Ay_CtaMonedaOrigen.XListCampo = "monedadescripcion"
      Ctr_Ay_CtaMonedaOrigen.Conexion VGCNx
    
    Case "B":
      Ctr_Ay_CtaMonedaOrigen.ListaCampos = "cbanco_codigo(1),cbanco_numero(1),monedasimbolo(1),cbanco_referenciacta(1),cbanco_nrocheque(1),monedacodigo(1)"
      Ctr_Ay_CtaMonedaOrigen.ListaCamposDescrip = "Código,Descripción,Mon,Ref,NCheque,MonCod"
      Ctr_Ay_CtaMonedaOrigen.ListaCamposText = "cbanco_codigo,cbanco_numero,monedasimbolo,cbanco_referenciacta,cbanco_nrocheque,monedacodigo"
      Ctr_Ay_CtaMonedaOrigen.NomTabla = "v_bancomoneda"
      Ctr_Ay_CtaMonedaOrigen.XcodCampo = "cbanco_codigo"
      Ctr_Ay_CtaMonedaOrigen.XListCampo = "cbanco_numero"
      Ctr_Ay_CtaMonedaOrigen.Conexion VGCNx
  End Select
End Sub

Sub ActivarCtaMonedaDestino()
  Select Case m_CasoDestino
    Case "C":
      Ctr_Ay_CtaMonedaDestino.ListaCampos = "monedacodigo(1),monedadescripcion(1)"
      Ctr_Ay_CtaMonedaDestino.ListaCamposDescrip = "Código,Descripción"
      Ctr_Ay_CtaMonedaDestino.ListaCamposText = "monedacodigo,monedadescripcion"
      Ctr_Ay_CtaMonedaDestino.NomTabla = "gr_moneda"
      Ctr_Ay_CtaMonedaDestino.XcodCampo = "monedacodigo"
      Ctr_Ay_CtaMonedaDestino.XListCampo = "monedadescripcion"
      Ctr_Ay_CtaMonedaDestino.Conexion VGCNx
      
    Case "B":
      Ctr_Ay_CtaMonedaDestino.ListaCampos = "cbanco_codigo(1),cbanco_numero(1),monedasimbolo(1),cbanco_referenciacta(1),monedacodigo(1)"
      Ctr_Ay_CtaMonedaDestino.ListaCamposDescrip = "Código,Descripción,Mon,Ref,MonCod"
      Ctr_Ay_CtaMonedaDestino.ListaCamposText = "cbanco_codigo,cbanco_numero,monedasimbolo,cbanco_referenciacta,monedacodigo"
      Ctr_Ay_CtaMonedaDestino.NomTabla = "v_bancomoneda"
      Ctr_Ay_CtaMonedaDestino.XcodCampo = "cbanco_codigo"
      Ctr_Ay_CtaMonedaDestino.XListCampo = "cbanco_numero"
      Ctr_Ay_CtaMonedaDestino.Conexion VGCNx
  End Select
End Sub
Sub ActivarConceptoTransf()
Call Ctr_Ay_Transf.Conexion(VGCNx)
   If m_cuentasxrendir = 1 Then
      Ctr_Ay_Transf.Filtro = " isnull(conceptoingresaref,0)=1"
    ElseIf m_fondofijo = 1 Then
           Ctr_Ay_Transf.Filtro = " isnull(conceptoingresaref,0)=1"
         Else
           Ctr_Ay_Transf.Filtro = "conceptocodigo like '" & Trim(VGParametros.transferenciaegreso) & "'"
 End If
End Sub
Sub ActualizaNumerador()
  Dim rb As New ADODB.Recordset
  Set rb = New ADODB.Recordset
    
    'Actualiza Numerador de Transferencia
    Set rb = VGCNx.Execute("select empresanumtransferencia from te_parametroempresa where empresacodigo='" & VGempresa & "'")
    If rb.RecordCount > 0 Then
       Text1(1).Text = Right("0000000000" & Trim(CStr(CDbl(IIf(IsNull(rb(0)) Or Len(Trim(rb(0))) = 0, 1, rb(0))))), 6)
       VGCNx.Execute "Update te_parametroempresa Set empresanumtransferencia='" & Right("0000000000" & Trim(CStr(Val(Text1(1) + 1))), 6) & "' where empresacodigo='" & VGempresa & "'"
    End If
    rb.Close
    Set rb = Nothing
  
    'Actualiza Numerador de Tipo de Egreso
    Set rb = VGCNx.Execute("select empresanumegreso from te_parametroempresa where empresacodigo='" & VGempresa & "'")
    If rb.RecordCount > 0 Then
       Text1(0).Text = Right("0000000000" & Trim(CStr(CDbl(IIf(IsNull(rb(0) + 1) Or Len(Trim(rb(0))) = 0, 1, rb(0) + 1)))), 6)
       VGCNx.Execute "Update te_parametroempresa Set empresanumegreso='" & Right("0000000000" & Trim(CStr(Val(Text1(0)))), 6) & "' where empresacodigo='" & VGempresa & "'"
    End If
    rb.Close
    Set rb = Nothing

End Sub

Function GrabarDataOrigen() As Integer
  Dim acmd As New ADODB.Command
  Dim xabono, xzona, xmone, xcuenta, xtipo As String
  Dim xnumplan, ximpsol, xtcam, xnumpag As Double
  
    GrabarDataOrigen = 0
    
    Set acmd.ActiveConnection = VGgeneral
    acmd.CommandType = adCmdStoredProc
    acmd.CommandText = "te_abonadocumento_pro"
    acmd.CommandTimeout = 0
    acmd.Prepared = True
    With acmd
        .Parameters("@base") = VGCNx.DefaultDatabase
        .Parameters("@tipo") = "1"
        .Parameters("@numrecibo") = Escadena(Text1(0).Text)
        .Parameters("@estadoreg") = ""
        .Parameters("@controlctacte") = ""
        .Parameters("@vendedorcodigo") = VGoficina
        .Parameters("@cajacodigo") = IIf(m_CasoOrigen = "C", Trim(Ctr_Ay_Origen.xclave), "")
        .Parameters("@clientecodigo") = IIf(RTrim(Ctr_AyuEntidad.xclave) = "", "", Ctr_AyuEntidad.xclave)
        .Parameters("@descripcion") = IIf(RTrim(Ctr_AyuEntidad.xclave) = "", "Transferencia A: ", Ctr_AyuEntidad.xnombre)
        .Parameters("@operacion") = VGParametros.codigooperaciontransferencia
        .Parameters("@monedacodigo") = lblMonOrigen.Caption
        .Parameters("@ingsal") = "E"
        .Parameters("@tipocambio") = Round(IIf(CDbl(numero(txt(1).Text)) = 0, 1#, CDbl(numero(txt(1).Text))), 4)
        .Parameters("@totsoles") = Round(IIf(lblMonOrigen.Caption = "01", CDbl(txt(2).Text), CDbl(txt(2).Text) * CDbl(txt(1).Text)), 4)
        .Parameters("@totdolares") = Round(IIf(lblMonOrigen.Caption = "01", CDbl(txt(2).Text) / CDbl(txt(1).Text), CDbl(txt(2).Text)), 4)
        .Parameters("@fechadocumento") = Format(DTPicker1.Value, "dd/mm/yyyy")
        .Parameters("@empresa") = Ctr_Ayuempresaorigen.xclave
        .Parameters("@observa") = txt(3).Text
        If cambio = 1 Then
           .Parameters("@transferauto") = ""
           .Parameters("@numreciboegreso") = Ctr_Ayutransf.xclave
         Else
           .Parameters("@transferauto") = "1"
           .Parameters("@numreciboegreso") = Text1(1).Text
        End If
        .Parameters("@usuario") = VGusuario
        .Parameters("@fechaact") = Date
        .Parameters("@saldodocxrendir") = IIf(lblMonOrigen.Caption = "01", CDbl(txt(2).Text), CDbl(txt(2).Text) * CDbl(txt(1).Text))
      End With
     acmd.Execute
     Set acmd = Nothing
     Set acmd.ActiveConnection = VGgeneral
     acmd.CommandType = adCmdStoredProc
     acmd.CommandText = "te_abonadetalledocumento_pro"
     acmd.CommandTimeout = 0
     acmd.Prepared = True
     With acmd
         .Parameters("@base") = VGCNx.DefaultDatabase
         .Parameters("@tipo") = "1"
         .Parameters("@numrecibo") = Text1(0).Text
         .Parameters("@estadoreg") = ""
         .Parameters("@item") = "1"
         .Parameters("@emisioncheque") = m_CasoOrigen   ' ver si es cheque
         .Parameters("@tipodocconcepto") = Ctr_Ay_Transf.xclave
         .Parameters("@numdocumento") = Escadena(txt(0).Text)
         .Parameters("@carabo") = "C"
         .Parameters("@formacan") = ""
         .Parameters("@tdqc") = Ctr_AyuDocumento.xclave
         .Parameters("@ndqc") = Escadena(txt(0).Text)
         .Parameters("@tipocajabanco") = m_CasoOrigen
         .Parameters("@cajabanco") = Ctr_Ay_Origen.xclave      'IIf(Len(Trim(Text1(2))) = 0, Trim(rsdetat.Fields(5)), Trim(Text1(2)))
         .Parameters("@numctacte") = IIf(m_CasoOrigen = "B", Trim(Ctr_Ay_CtaMonedaOrigen.xnombre), "")  'numero de cuenta corriente
         .Parameters("@adicionactacte") = ""
         .Parameters("@monedadocumento") = ""
         .Parameters("@monedacancela") = lblMonOrigen.Caption
         .Parameters("@importesoles") = Round(IIf(lblMonOrigen.Caption = "01", CDbl(txt(2).Text), CDbl(txt(2).Text) * CDbl(txt(1).Text)), 4)
         .Parameters("@importedolares") = Round(IIf(lblMonOrigen.Caption = "01", CDbl(txt(2).Text) / CDbl(txt(1).Text), CDbl(txt(2).Text)), 4)
         .Parameters("@contabledisponi") = "" 'Escadena(VGParametros.saldocontadispo)      'sale de empresas
         .Parameters("@fechacancela") = Format(DTPicker1.Value, "dd/mm/yyyy")
         .Parameters("@observacion") = txt(3).Text
         .Parameters("@usuario") = VGusuario
         .Parameters("@fechaact") = Date
     End With
     acmd.Execute
     Set acmd = Nothing
     DoEvents
     GrabarDataOrigen = 1
End Function

Sub NumeradorIngreso()
 Dim rb As New ADODB.Recordset
    'Actualiza el Numerador de Tipo de Ingreso
    
    Set rb = New ADODB.Recordset
    Set rb = VGCNx.Execute("select empresanumeingreso from te_parametroempresa where empresacodigo='" & VGempresa & "'")
    If rb.RecordCount > 0 Then
       Text1(2).Text = Right("0000000000" & Trim(CStr(CDbl(IIf(IsNull(rb(0) + 1) Or Len(Trim(rb(0))) = 0, 1, rb(0) + 1)))), 6)
       VGCNx.Execute "Update te_parametroempresa Set empresanumeingreso='" & Right("0000000000" & Trim(CStr(Val(Text1(2)))), 6) & "' where empresacodigo='" & VGempresa & "'"
    End If
    rb.Close
    Set rb = Nothing

End Sub

Function GrabarDataDestino() As Integer
  Dim acmd As New ADODB.Command
  Dim rb As New ADODB.Recordset
  Dim xabono, xzona, xmone, xcuenta, xtipo As String
  Dim xnumplan, ximpsol, xtcam, xnumpag As Double
  
    GrabarDataDestino = 0
    
    Set acmd.ActiveConnection = VGgeneral
    acmd.CommandType = adCmdStoredProc
    acmd.CommandText = "te_abonadocumento_pro"
    acmd.CommandTimeout = 0
    acmd.Prepared = True
    With acmd
        .Parameters("@base") = VGCNx.DefaultDatabase
        .Parameters("@tipo") = "1"
        .Parameters("@numrecibo") = Escadena(Text1(2).Text)
        .Parameters("@estadoreg") = ""
        .Parameters("@controlctacte") = ""
        .Parameters("@vendedorcodigo") = VGoficina
        .Parameters("@cajacodigo") = IIf(m_CasoDestino = "C", Trim(Ctr_Ay_Destino.xclave), "")
        .Parameters("@clientecodigo") = IIf(RTrim(Ctr_AyuEntidad.xclave) = "", "", Ctr_AyuEntidad.xclave)
        .Parameters("@descripcion") = IIf(RTrim(Ctr_AyuEntidad.xclave) = "", "Transferencia De : ", Ctr_AyuEntidad.xnombre)
        .Parameters("@operacion") = VGParametros.codigooperaciontransferencia
        .Parameters("@monedacodigo") = lblMonDes.Caption
        .Parameters("@ingsal") = "I"
        .Parameters("@tipocambio") = Round(CDbl(txt(1).Text), 4)
        .Parameters("@totsoles") = Round(IIf(lblMonOrigen.Caption = "01", CDbl(txt(2).Text), CDbl(txt(2).Text) * CDbl(txt(1).Text)), 4)
        .Parameters("@totdolares") = Round(IIf(lblMonOrigen.Caption = "01", CDbl(txt(2).Text) / CDbl(txt(1).Text), CDbl(txt(2).Text)), 4)
        .Parameters("@fechadocumento") = Format(DTPicker1.Value, "dd/mm/yyyy")
        .Parameters("@observa") = txt(3).Text
        .Parameters("@empresa") = Ctr_Ayuempresadestino.xclave
        If cambio = 1 Then
           .Parameters("@transferauto") = ""
           .Parameters("@numreciboegreso") = Ctr_Ayutransf.xclave
         Else
           .Parameters("@transferauto") = "1"
           .Parameters("@numreciboegreso") = Text1(1).Text
        End If
        .Parameters("@usuario") = VGusuario
        .Parameters("@fechaact") = Date
        .Parameters("@saldodocxrendir") = IIf(lblMonOrigen.Caption = "01", CDbl(txt(2).Text), CDbl(txt(2).Text) * CDbl(txt(1).Text))
     End With
     acmd.Execute
     Set acmd = Nothing
     Set acmd.ActiveConnection = VGgeneral
     acmd.CommandType = adCmdStoredProc
     acmd.CommandText = "te_abonadetalledocumento_pro"
     acmd.CommandTimeout = 0
     acmd.Prepared = True
     With acmd
         .Parameters("@base") = VGCNx.DefaultDatabase
         .Parameters("@tipo") = "1"
         .Parameters("@numrecibo") = Escadena(Text1(2).Text)
         .Parameters("@estadoreg") = ""
         .Parameters("@item") = "1"
         .Parameters("@emisioncheque") = m_CasoDestino     ' ver si es cheque
         .Parameters("@tipodocconcepto") = Ctr_Ay_Transf.xclave
         .Parameters("@numdocumento") = Trim(txt(0).Text)
         .Parameters("@carabo") = "C"
         .Parameters("@formacan") = ""
         .Parameters("@tdqc") = Ctr_AyuDocumento.xclave
         .Parameters("@ndqc") = Escadena(txt(0).Text)
         .Parameters("@tipocajabanco") = m_CasoDestino
         .Parameters("@cajabanco") = Trim(Ctr_Ay_Destino.xclave)
         .Parameters("@numctacte") = Trim(IIf(m_CasoDestino = "B", Trim(Ctr_Ay_CtaMonedaDestino.xnombre), "")) 'numero de cuenta corriente
         .Parameters("@adicionactacte") = ""
         .Parameters("@monedadocumento") = ""
         .Parameters("@monedacancela") = lblMonDes.Caption
         .Parameters("@importesoles") = Round(IIf(lblMonOrigen.Caption = "01", CDbl(txt(2).Text), CDbl(txt(2).Text) * CDbl(txt(1).Text)), 4)  'CDbl(IIf(rsdetat.Fields(7) = g_TipoSol, rsdetat.Fields(8), (rsdetat.Fields(8) * xtcam)))
         .Parameters("@importedolares") = Round(IIf(lblMonOrigen.Caption = "01", CDbl(txt(2).Text) / CDbl(txt(1).Text), CDbl(txt(2).Text)), 4) 'CDbl(IIf(rsdetat.Fields(7) = g_TipoSol, (rsdetat.Fields(8) / xtcam), rsdetat.Fields(8)))
         .Parameters("@contabledisponi") = "" 'Escadena(VGParametros.saldocontadispo)      'sale de empresas
         .Parameters("@fechacancela") = Format(DTPicker1.Value, "dd/mm/yyyy")
         .Parameters("@observacion") = txt(3).Text
         .Parameters("@usuario") = VGusuario
         .Parameters("@fechaact") = Date
     End With
     acmd.Execute
     Set acmd = Nothing
     DoEvents
     GrabarDataDestino = 1
     MsgBox "Los datos han sido grabados satisfactoriamente...!!!", vbInformation, MsgTitle
End Function

Private Sub ImpresionTransferencias(xNumTrans As String)
 Dim rs As New ADODB.Recordset
 Dim SQL As String
   Set rs = New ADODB.Recordset
   SQL = "Select cabrec_numrecibo from te_cabecerarecibos "
   SQL = SQL & "Where cabrec_numreciboegreso<>'' and cabrec_numreciboegreso='" & Text1(1).Text & "' "
   SQL = SQL & "and cabrec_estadoreg<>'1'"
  
   Set rs = VGCNx.Execute(SQL)
   If Not rs.BOF And Not rs.EOF Then
      rs.MoveFirst
      Do Until rs.EOF
         Call ImprimirRecibo(rs(0))
         rs.MoveNext
      Loop
   End If

End Sub

Function ValidarData() As Boolean
 Dim rs As ADODB.Recordset
 Dim SQL As String
  If Trim(Ctr_Ay_CtaMonedaOrigen.xclave) = Trim(Ctr_Ay_CtaMonedaDestino.xclave) And _
    Trim(Ctr_Ay_CtaMonedaOrigen.xnombre) = Trim(Ctr_Ay_CtaMonedaDestino.xnombre) And _
      m_CasoOrigen = m_CasoDestino And m_CasoOrigen = "B" Then
       MsgBox "La Cuenta de Banco Destino no puede ser la misma del Banco Origen"
       Ctr_Ay_CtaMonedaDestino.xclave = Empty
       Ctr_Ay_CtaMonedaDestino.xnombre = Empty
       Ctr_Ay_CtaMonedaDestino.SetFocus
       ValidarData = False
       Exit Function
  End If
        
  If m_CasoOrigen = "B" Then
     Set rs = New ADODB.Recordset
     SQL = "Select count(*) from te_detallerecibos where "
     SQL = SQL & ""
     SQL = "select count(detrec_numctacte) from te_detallerecibos "
     SQL = SQL & "where detrec_tipocajabanco like 'B' and detrec_cajabanco1='" & Trim(Ctr_Ay_Origen.xclave) & "' and "
     SQL = SQL & "detrec_monedacancela='" & lblMonOrigen.Caption & "' AND "
     SQL = SQL & "detrec_numctacte='" & Trim(txt(0).Text) & "'"
     Set rs = VGCNx.Execute(SQL)
     If rs(0) > 0 Then
        MsgBox "El Nº de Cheque: " & Trim(txt(0).Text) & " del Banco Seleccionado Existe", vbInformation, Caption
        ValidarData = False
        txt(0).SetFocus
        Exit Function
     End If
  End If
  
  If Ctr_Ay_Origen.xclave = Empty Then
     MsgBox "Falta Completar el Origen", vbInformation, Caption
     ValidarData = False
     Ctr_Ay_Origen.SetFocus
     Exit Function
  End If
  
  If Ctr_Ay_Destino.xclave = Empty Then
     MsgBox "Falta Completar el Destino", vbInformation, Caption
     ValidarData = False
     Ctr_Ay_Destino.SetFocus
     Exit Function
  End If
  
  If txt(1).Text = Empty And Trim(Ctr_Ay_CtaMonedaOrigen.xclave) <> Trim(Ctr_Ay_CtaMonedaDestino.xclave) Then
     MsgBox "Falta Completar el Tipo de Cambio", vbInformation, Caption
     ValidarData = False
     txt(1).SetFocus
     Exit Function
   ElseIf txt(1).Text = Empty Then
          txt(1).Text = 1#
     Else
       txt(1).Text = numero(txt(1).Text)
  End If
  If txt(2).Text = Empty Then
     MsgBox "Falta Completar el Importe a Transferir", vbInformation, Caption
     ValidarData = False
     txt(2).SetFocus
     Exit Function
  End If
        
  ValidarData = True
End Function

Sub MuestraNumeradorTransf()
  Dim rb As ADODB.Recordset
    Set rb = New ADODB.Recordset
    Set rb = VGCNx.Execute("select * from te_parametroempresa where empresacodigo='" & VGempresa & "'")
    If rb.RecordCount > 0 Then
       Text1(1).Text = Right("0000000000" & Trim(CStr(CDbl(IIf(IsNull(rb!empresanumtransferencia) Or Len(Trim(rb!empresanumtransferencia)) = 0, 1, rb!empresanumtransferencia)))), 6)
       Text1(0).Text = Right("0000000000" & Trim(CStr(CDbl(IIf(IsNull(rb!empresanumtransferencia) Or Len(Trim(rb!empresanumegreso)) = 0, 1, rb!empresanumegreso)))), 6)
       Text1(2).Text = Right("0000000000" & Trim(CStr(CDbl(IIf(IsNull(rb!empresanumtransferencia) Or Len(Trim(rb!empresanumeingreso)) = 0, 1, rb!empresanumeingreso)))), 6)
    
    End If
    rb.Close
    Set rb = Nothing
End Sub

