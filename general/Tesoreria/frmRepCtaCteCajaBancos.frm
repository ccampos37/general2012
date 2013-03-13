VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Begin VB.Form frmRepCtaCteCajaBancos 
   Caption         =   "Reporte de Cuenta Corriente"
   ClientHeight    =   4560
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6645
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4560
   ScaleWidth      =   6645
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   675
      Left            =   195
      TabIndex        =   17
      Top             =   195
      Width           =   6210
      Begin VB.OptionButton opt 
         Caption         =   "Caja"
         Height          =   300
         Index           =   0
         Left            =   1350
         TabIndex        =   0
         Top             =   270
         Width           =   1440
      End
      Begin VB.OptionButton opt 
         Caption         =   "Banco"
         Height          =   300
         Index           =   1
         Left            =   3360
         TabIndex        =   1
         Top             =   240
         Width           =   1440
      End
   End
   Begin VB.Frame fraDetallado 
      Height          =   2670
      Left            =   180
      TabIndex        =   10
      Top             =   885
      Width           =   6225
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyudaBancoCuenta 
         Height          =   315
         Left            =   1140
         TabIndex        =   7
         Top             =   2220
         Width           =   4845
         _ExtentX        =   8546
         _ExtentY        =   556
         XcodMaxLongitud =   20
         xcodwith        =   1400
         NomTabla        =   "ct_subasiento"
         ListaCampos     =   "subasientocodigo(1),subasientodescripcion(1)"
         XcodCampo       =   "subasientocodigo"
         XListCampo      =   "subasientodescripcion"
         ListaCamposDescrip=   "Código,Descripción"
         ListaCamposText =   "subasientocodigo,subasientodescripcion"
         Requerido       =   0   'False
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyudaBanco 
         Height          =   300
         Left            =   1140
         TabIndex        =   6
         Top             =   1905
         Width           =   4860
         _ExtentX        =   8573
         _ExtentY        =   529
         XcodMaxLongitud =   3
         xcodwith        =   800
         NomTabla        =   "ct_asiento"
         ListaCampos     =   "asientocodigo(1),asientodescripcion(1)"
         XcodCampo       =   "asientocodigo"
         XListCampo      =   "asientodescripcion"
         ListaCamposDescrip=   "Código,Descripción"
         ListaCamposText =   "asientocodigo,asientodescripcion"
         Requerido       =   0   'False
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayu_Caja 
         Height          =   360
         Left            =   1140
         TabIndex        =   4
         Top             =   1080
         Width           =   4845
         _ExtentX        =   8546
         _ExtentY        =   635
         XcodMaxLongitud =   0
         xcodwith        =   800
         NomTabla        =   "ct_cuenta"
         ListaCampos     =   "cuentacodigo(1),cuentadescripcion(1)"
         XcodCampo       =   "cuentacodigo"
         XListCampo      =   "cuentadescripcion"
         ListaCamposDescrip=   "Código,Descripción"
         ListaCamposText =   "cuentacodigo,cuentadescripcion"
         Requerido       =   0   'False
      End
      Begin MSComCtl2.DTPicker DTPickerFecFinal 
         Height          =   300
         Left            =   4125
         TabIndex        =   3
         Top             =   750
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   529
         _Version        =   393216
         Format          =   60817409
         CurrentDate     =   37474
      End
      Begin MSComCtl2.DTPicker DTPickerFecInicio 
         Height          =   300
         Left            =   1140
         TabIndex        =   2
         Top             =   750
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   529
         _Version        =   393216
         Format          =   60817409
         CurrentDate     =   37474
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayu_Moneda 
         Height          =   360
         Left            =   1125
         TabIndex        =   5
         Top             =   1440
         Width           =   4860
         _ExtentX        =   8573
         _ExtentY        =   635
         XcodMaxLongitud =   0
         xcodwith        =   800
         NomTabla        =   "ct_cuenta"
         ListaCampos     =   "cuentacodigo(1),cuentadescripcion(1)"
         XcodCampo       =   "cuentacodigo"
         XListCampo      =   "cuentadescripcion"
         ListaCamposDescrip=   "Código,Descripción"
         ListaCamposText =   "cuentacodigo,cuentadescripcion"
         Requerido       =   0   'False
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuempresa 
         Height          =   315
         Left            =   1125
         TabIndex        =   18
         Top             =   360
         Width           =   4695
         _ExtentX        =   8281
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
         Requerido       =   0   'False
      End
      Begin VB.Label Lblempresa 
         AutoSize        =   -1  'True
         Caption         =   "Empresa :"
         Height          =   195
         Left            =   240
         TabIndex        =   19
         Top             =   420
         Width           =   705
      End
      Begin VB.Label Label6 
         Caption         =   "Moneda"
         Height          =   225
         Left            =   150
         TabIndex        =   16
         Top             =   1500
         Width           =   885
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha Final"
         Height          =   285
         Left            =   3225
         TabIndex        =   15
         Top             =   795
         Width           =   840
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Inicial"
         Height          =   300
         Left            =   75
         TabIndex        =   14
         Top             =   780
         Width           =   930
      End
      Begin VB.Label Label3 
         Caption         =   "Caja"
         Height          =   255
         Left            =   150
         TabIndex        =   13
         Top             =   1170
         Width           =   885
      End
      Begin VB.Label Label4 
         Caption         =   "Banco"
         Height          =   255
         Left            =   180
         TabIndex        =   12
         Top             =   1950
         Width           =   885
      End
      Begin VB.Label Label5 
         Caption         =   "Cuenta"
         Height          =   285
         Left            =   150
         TabIndex        =   11
         Top             =   2295
         Width           =   885
      End
   End
   Begin VB.CommandButton cmdBotones 
      Caption         =   "&Cancelar"
      Height          =   360
      Index           =   1
      Left            =   3375
      TabIndex        =   9
      Top             =   3975
      Width           =   1215
   End
   Begin VB.CommandButton cmdBotones 
      Caption         =   "&Aceptar"
      Height          =   360
      Index           =   0
      Left            =   2025
      TabIndex        =   8
      Top             =   3975
      Width           =   1215
   End
End
Attribute VB_Name = "frmRepCtaCteCajaBancos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Ctr_AyudaBanco_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
  Ctr_AyudaBancoCuenta.Filtro = "cbanco_codigo='" & ColecCampos("bancocodigo").Value & "'"
End Sub

Private Sub Form_Load()
  Dim cFecha As Date
  Opt(0).Value = True
  Call Ctr_Ayuempresa.Conexion(VGCNx)
  DTPickerFecInicio.Value = Fecha(1, VGParamSistem.fechatrabajo)
  DTPickerFecFinal.Value = Fecha(2, VGParamSistem.fechatrabajo)
  End Sub

Private Sub cmdBotones_Click(Index As Integer)
  Select Case Index
    Case 0:
      Call ImpresionEstadoCtaCte
    Case 1:
      Unload Me
  End Select
End Sub

Sub ImpresionEstadoCtaCte()
Dim arrform() As Variant, arrparm() As Variant
    ReDim arrparm(10)
    ReDim arrform(2)
    arrparm(0) = VGParamSistem.BDEmpresa
    arrparm(1) = IIf(Opt(0).Value = True, "C", "B")
    If Opt(0).Value = True Then
       arrparm(2) = IIf(Ctr_Ayu_Caja.xclave = Empty, "%", Trim(Ctr_Ayu_Caja.xclave))
       arrparm(3) = IIf(Ctr_Ayu_Moneda.xclave = Empty, "%", Trim(Ctr_Ayu_Moneda.xclave))
    Else
       arrparm(2) = IIf(Ctr_AyudaBanco.xclave = Empty, "%", Trim(Ctr_AyudaBanco.xclave))
       arrparm(3) = IIf(Ctr_AyudaBancoCuenta.xclave = Empty, "%", Trim(Ctr_AyudaBancoCuenta.xclave))
    End If
    arrparm(4) = Format(DTPickerFecInicio.Value, "dd/mm/yyyy")
    arrparm(5) = Format(DTPickerFecFinal.Value, "dd/mm/yyyy")
    arrparm(6) = IIf(Ctr_Ayuempresa.xclave = Empty, "%%", Trim(Ctr_Ayuempresa.xclave))
    arrparm(7) = "%%"
    arrparm(8) = "%%"
    arrparm(9) = "%%"
    arrform(0) = "@Empresa='" & VGParametros.descripcion & "'"
    Call ImpresionRptProc("te_CtaCtexCajaBanco.rpt", arrform, arrparm, "Saldos de Caja Bancos")
End Sub

Sub ConfiguraCajaBanco(valor As Boolean)
  Ctr_Ayu_Caja.Enabled = valor
  Ctr_Ayu_Moneda.Enabled = valor
  Ctr_AyudaBanco.Enabled = Not valor
  Ctr_AyudaBancoCuenta.Enabled = Not valor
  If valor = True Then
     Ctr_Ayu_Caja.ListaCampos = "cajacodigo(1),cajadescripcion(1)"
     Ctr_Ayu_Caja.ListaCamposDescrip = "Código,Descripción"
     Ctr_Ayu_Caja.ListaCamposText = "cajacodigo,cajadescripcion"
     Ctr_Ayu_Caja.NomTabla = "te_codigocaja"
     Ctr_Ayu_Caja.XcodCampo = "cajacodigo"
     Ctr_Ayu_Caja.XListCampo = "cajadescripcion"
     Ctr_Ayu_Caja.Conexion VGCNx
  Else
     Ctr_AyudaBanco.ListaCampos = "bancocodigo(1),bancodescripcion(1)"
     Ctr_AyudaBanco.ListaCamposDescrip = "Código,Descripción"
     Ctr_AyudaBanco.ListaCamposText = "bancocodigo,bancodescripcion"
     Ctr_AyudaBanco.NomTabla = "gr_banco"
     Ctr_AyudaBanco.XcodCampo = "bancocodigo"
     Ctr_AyudaBanco.XListCampo = "bancodescripcion"
     Ctr_AyudaBanco.Conexion VGCNx
  End If
  
  If valor = True Then
      Ctr_Ayu_Moneda.ListaCampos = "monedacodigo(1),monedadescripcion(1)"
      Ctr_Ayu_Moneda.ListaCamposDescrip = "Código,Descripción"
      Ctr_Ayu_Moneda.ListaCamposText = "monedacodigo,monedadescripcion"
      Ctr_Ayu_Moneda.NomTabla = "gr_moneda"
      Ctr_Ayu_Moneda.XcodCampo = "monedacodigo"
      Ctr_Ayu_Moneda.XListCampo = "monedadescripcion"
      Ctr_Ayu_Moneda.Conexion VGCNx
  Else
      Ctr_AyudaBancoCuenta.ListaCampos = "cbanco_codigo(1),cbanco_numero(1),cbanco_referenciacta(1),cbanco_nrocheque(1),monedacodigo(1)"
      Ctr_AyudaBancoCuenta.ListaCamposDescrip = "Código,Descripción,Mon,Ref,NCheque,MonCod"
      Ctr_AyudaBancoCuenta.ListaCamposText = "cbanco_codigo,cbanco_numero,cbanco_referenciacta,cbanco_nrocheque,monedacodigo"
      Ctr_AyudaBancoCuenta.NomTabla = "te_cuentabancos"
      Ctr_AyudaBancoCuenta.XcodCampo = "cbanco_numero"
      Ctr_AyudaBancoCuenta.XListCampo = "cbanco_referenciacta"
      Ctr_AyudaBancoCuenta.Conexion VGCNx
  End If
  
End Sub

Sub ConfiguraBanco(valor As Boolean)
  Ctr_AyudaBanco.Enabled = valor
  Ctr_AyudaBancoCuenta.Enabled = valor
End Sub

Private Sub Opt_Click(Index As Integer)
  Select Case Index
    Case 0:
       Call ConfiguraCajaBanco(True)
    
    Case 1:
       Call ConfiguraCajaBanco(False)
  End Select

End Sub
