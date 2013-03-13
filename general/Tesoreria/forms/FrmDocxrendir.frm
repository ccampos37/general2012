VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Begin VB.Form FrmDocxrendir 
   Caption         =   "Documentos x rendir"
   ClientHeight    =   5070
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6915
   LinkTopic       =   "Form1"
   ScaleHeight     =   5070
   ScaleWidth      =   6915
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame6 
      Height          =   975
      Left            =   3240
      TabIndex        =   3
      Top             =   3885
      Width           =   3255
      Begin VB.CommandButton Cmdbotones 
         Caption         =   "&Aceptar"
         Height          =   660
         Index           =   0
         Left            =   240
         Picture         =   "FrmDocxrendir.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   200
         Width           =   1035
      End
      Begin VB.CommandButton Cmdbotones 
         Caption         =   "&Salir"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Index           =   1
         Left            =   1800
         Picture         =   "FrmDocxrendir.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   200
         Width           =   1035
      End
   End
   Begin VB.Frame FrameRendicion 
      Caption         =   "Rendicion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   240
      TabIndex        =   28
      Top             =   1005
      Width           =   6255
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyudaOficina 
         Height          =   300
         Left            =   1050
         TabIndex        =   29
         Top             =   240
         Width           =   5100
         _ExtentX        =   8996
         _ExtentY        =   529
         XcodMaxLongitud =   3
         xcodwith        =   400
         NomTabla        =   "cp_oficina"
         TituloAyuda     =   "Ayuda de Caja"
         ListaCampos     =   "vendedorcodigo(1),vendedornombres(1)"
         XcodCampo       =   "vendedorcodigo"
         XListCampo      =   "vendedornombres"
         ListaCamposDescrip=   "Código,Descripción"
         ListaCamposText =   "vendedorcodigo,vendedornombres"
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyuRendicion1 
         Height          =   315
         Left            =   908
         TabIndex        =   30
         Top             =   1080
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         XcodMaxLongitud =   0
         xcodwith        =   800
         NomTabla        =   "te_rendiciones"
         TituloAyuda     =   "Busqueda de Rendiciones"
         ListaCampos     =   "rendicionnumero(1),monedacodigo(1),rendicionfecha(2)"
         XcodCampo       =   "rendicionnumero"
         XListCampo      =   "monedacodigo"
         ListaCamposDescrip=   "Nro Rendicion,Moneda, fecha rendicion"
         ListaCamposText =   "rendicionnumero,monedacodigo,rendicionfecha"
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyudaCaja 
         Height          =   315
         Left            =   1050
         TabIndex        =   31
         Top             =   690
         Width           =   5100
         _ExtentX        =   8996
         _ExtentY        =   556
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
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayurendicion2 
         Height          =   315
         Left            =   3908
         TabIndex        =   32
         Top             =   1080
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         XcodMaxLongitud =   0
         xcodwith        =   800
         NomTabla        =   "te_rendiciones"
         TituloAyuda     =   "Busqueda de Rendiciones"
         ListaCampos     =   "rendicionnumero(1),monedacodigo(1),rendicionfecha(2)"
         XcodCampo       =   "rendicionnumero"
         XListCampo      =   "monedacodigo"
         ListaCamposDescrip=   "Nro Rendicion,Moneda, fecha rendicion"
         ListaCamposText =   "rendicionnumero,monedacodigo,rendicionfecha"
      End
      Begin VB.Label Label4 
         Caption         =   "Oficina"
         Height          =   255
         Index           =   0
         Left            =   150
         TabIndex        =   36
         Top             =   285
         Width           =   885
      End
      Begin VB.Label lbMon 
         Caption         =   "Desde :"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   1095
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Cod. Caja"
         Height          =   255
         Index           =   1
         Left            =   135
         TabIndex        =   34
         Top             =   690
         Width           =   885
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta :"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   3240
         TabIndex        =   33
         Top             =   1125
         Width           =   735
      End
   End
   Begin VB.Frame fraDetallado 
      Caption         =   "Caja Bancos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1470
      Left            =   240
      TabIndex        =   15
      Top             =   1005
      Width           =   6285
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyudaBancoCuenta 
         Height          =   315
         Left            =   1140
         TabIndex        =   16
         Top             =   930
         Width           =   4845
         _ExtentX        =   8546
         _ExtentY        =   556
         XcodMaxLongitud =   4
         xcodwith        =   800
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
         TabIndex        =   17
         Top             =   615
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
         TabIndex        =   18
         Top             =   600
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
         TabIndex        =   19
         Top             =   270
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   529
         _Version        =   393216
         Format          =   57802753
         CurrentDate     =   37474
      End
      Begin MSComCtl2.DTPicker DTPickerFecInicio 
         Height          =   300
         Left            =   1140
         TabIndex        =   20
         Top             =   270
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   529
         _Version        =   393216
         Format          =   57802753
         CurrentDate     =   37474
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayu_Moneda 
         Height          =   360
         Left            =   1125
         TabIndex        =   21
         Top             =   960
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
      Begin VB.Label lmon 
         Caption         =   "Moneda"
         Height          =   225
         Left            =   150
         TabIndex        =   27
         Top             =   1005
         Width           =   885
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha Final"
         Height          =   285
         Left            =   3225
         TabIndex        =   26
         Top             =   315
         Width           =   840
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Inicial"
         Height          =   300
         Left            =   75
         TabIndex        =   25
         Top             =   285
         Width           =   930
      End
      Begin VB.Label lcaja 
         Caption         =   "Caja"
         Height          =   255
         Left            =   150
         TabIndex        =   24
         Top             =   660
         Width           =   885
      End
      Begin VB.Label lban 
         Caption         =   "Banco"
         Height          =   255
         Left            =   150
         TabIndex        =   23
         Top             =   660
         Width           =   885
      End
      Begin VB.Label lcta 
         Caption         =   "Cuenta"
         Height          =   285
         Left            =   150
         TabIndex        =   22
         Top             =   1005
         Width           =   885
      End
   End
   Begin VB.Frame FrameCajaBancos 
      Height          =   795
      Left            =   3915
      TabIndex        =   12
      Top             =   120
      Width           =   2610
      Begin VB.OptionButton Opt 
         Caption         =   "Banco"
         Height          =   300
         Index           =   1
         Left            =   1320
         TabIndex        =   14
         Top             =   360
         Width           =   960
      End
      Begin VB.OptionButton Opt 
         Caption         =   "Caja"
         Height          =   300
         Index           =   0
         Left            =   270
         TabIndex        =   13
         Top             =   360
         Width           =   960
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
      Height          =   1215
      Left            =   240
      TabIndex        =   7
      Top             =   2565
      Width           =   6285
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_concepto1 
         Height          =   315
         Left            =   1230
         TabIndex        =   8
         Top             =   360
         Width           =   4845
         _ExtentX        =   8546
         _ExtentY        =   556
         XcodMaxLongitud =   4
         xcodwith        =   800
         NomTabla        =   "te_conceptocaja"
         TituloAyuda     =   "Ayuda de Conceptos"
         ListaCampos     =   "conceptocodigo(1),conceptodescripcion(1)"
         XcodCampo       =   "conceptocodigo"
         XListCampo      =   "conceptodescripcion"
         ListaCamposDescrip=   "Código,Descripción"
         ListaCamposText =   "conceptocodigo,conceptodescripcion"
         Requerido       =   0   'False
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuempresa 
         Height          =   315
         Left            =   1230
         TabIndex        =   9
         Top             =   750
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
         Left            =   345
         TabIndex        =   11
         Top             =   810
         Width           =   705
      End
      Begin VB.Label Label7 
         Caption         =   "Concepto"
         Height          =   285
         Left            =   255
         TabIndex        =   10
         Top             =   420
         Width           =   885
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Imprimir "
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
      Height          =   735
      Left            =   240
      TabIndex        =   4
      Top             =   165
      Width           =   3495
      Begin VB.OptionButton OptRendiciones 
         Caption         =   "Nro de Rendicion"
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton OptCajaBancos 
         Caption         =   "Caja Bancos"
         Height          =   375
         Left            =   2160
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo"
      ForeColor       =   &H00FF0000&
      Height          =   975
      Left            =   240
      TabIndex        =   0
      Top             =   3885
      Width           =   2895
      Begin VB.OptionButton OptDetallado 
         Caption         =   "Detallado"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   1095
      End
      Begin VB.OptionButton OptResumido 
         Caption         =   "Resumido"
         Height          =   195
         Left            =   1440
         TabIndex        =   1
         Top             =   480
         Width           =   1095
      End
   End
End
Attribute VB_Name = "FrmDocxrendir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim valorop As String
Dim valoroptext As String
Private Sub Ctr_AyudaBanco_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
  Ctr_AyudaBancoCuenta.Filtro = "cbanco_codigo='" & ColecCampos("bancocodigo").Value & "'"
End Sub

Private Sub Ctr_AyudaCaja_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
 Ctr_AyuRendicion1.Filtro = " codigocaja='" & Ctr_AyudaCaja.xclave & "'"
 Ctr_Ayurendicion2.Filtro = " codigocaja='" & Ctr_AyudaCaja.xclave & "'"
 End Sub

Private Sub Form_Load()
  Dim cFecha As Date
  Dim rsql As New ADODB.Recordset
  Set rsql = VGcnx.Execute("select * from te_codigocaja where cajafondofijo=1 or cajacuentaxrendir=1")
If rsql.RecordCount = 0 Then
  If MsgBox("No existe registros para este tipo de reporte, Desea cancelar ", vbQuestion + vbOKCancel) = vbOK Then
    Cmdbotones(0).Enabled = False
    Exit Sub
  End If
  Else
  Opt(0).Value = True
  Me.Width = 6860
  Me.Height = 6795
  DTPickerFecInicio.Value = Format("01/" & Format(Month(VGParamSistem.fechatrabajo), "00") & "/" & Year(VGParamSistem.fechatrabajo), "dd/mm/yyyy")
  cFecha = Format("01/" & Format(Month(VGParamSistem.fechatrabajo) + 1, "00") & "/" & Year(VGParamSistem.fechatrabajo), "dd/mm/yyyy")
  DTPickerFecFinal.Value = Format(cFecha - 1, "dd/mm/yyyy")
  Call Ctr_concepto1.Conexion(VGcnx)
  Call Ctr_Ayuempresa.Conexion(VGcnx)
  Call Ctr_AyuRendicion1.Conexion(VGcnx)
  Call Ctr_Ayurendicion2.Conexion(VGcnx)
  Call Ctr_AyudaOficina.Conexion(VGcnx)
  Call Ctr_AyudaCaja.Conexion(VGcnx)
  Ctr_Ayu_Caja.Filtro = " cajafondofijo=1 or cajacuentaxrendir=1 "
 OptCajaBancos.Value = True
  OptDetallado.Value = True
 End If
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
    ReDim arrparm(12)
    ReDim arrform(6)
valorop = "%%"
    arrparm(0) = VGParamSistem.BDEmpresa
    If OptCajaBancos.Value = True Then
       arrparm(1) = IIf(Opt(0).Value = True, "C", "B")
       If Opt(0).Value = True Then
          arrparm(2) = IIf(Ctr_Ayu_Caja.xclave = Empty, "%%", Trim(Ctr_Ayu_Caja.xclave))
          arrparm(3) = Trim(IIf(Ctr_Ayu_Moneda.xnombre = Empty, "%%", Trim(Ctr_Ayu_Moneda.xclave)))
        Else
          arrparm(2) = IIf(Ctr_AyudaBanco.xclave = Empty, "%%", Trim(Ctr_AyudaBanco.xclave))
          arrparm(3) = Trim(IIf(Ctr_AyudaBancoCuenta.xnombre = Empty, "%%", Trim(Ctr_AyudaBancoCuenta.xnombre)))
       End If
       arrparm(4) = Format(DTPickerFecInicio.Value, "dd/mm/yyyy")
       arrparm(5) = Format(DTPickerFecFinal.Value, "dd/mm/yyyy")
       arrform(4) = "rangofecha=' DEL : " & Format(DTPickerFecInicio.Value, "dd/mm/yyyy") & " AL " & Format(DTPickerFecFinal.Value, "dd/mm/yyyy") & "'"
     Else
       arrparm(1) = "%"
       arrparm(2) = IIf(Ctr_AyudaCaja.xclave = Empty, "%%", Trim(Ctr_AyudaCaja.xclave))
       arrparm(3) = "%%"
       arrparm(4) = Ctr_AyuRendicion1.xclave
       arrparm(5) = Ctr_Ayurendicion2.xclave
       arrform(4) = "rangofecha=' DESDE RENDICION :  " & Ctr_AyuRendicion1.xclave & "  HASTA :  " & Ctr_Ayurendicion2.xclave & "'"
 End If
 arrparm(6) = IIf(Trim(Ctr_concepto1.xclave) = "", "%%", Trim(Ctr_concepto1.xclave))
 arrparm(7) = "%%"
 arrparm(8) = valorop
 arrparm(9) = IIf(Ctr_Ayuempresa.xclave = Empty, "%%", Trim(Ctr_Ayuempresa.xclave))
 arrparm(10) = IIf(OptRendiciones.Value = True, "0", "1")
 arrparm(11) = IIf(OptDetallado.Value = True, "0", "1")
 
 arrform(0) = "@Empresa='" & VGParametros.descripcion & "'"
 arrform(1) = "concep='" & IIf(Trim(Ctr_concepto1.xnombre) = "", "Todos", Trim(Ctr_concepto1.xnombre)) & "'"
 arrform(2) = ""
 arrform(3) = "transfer=''"
 If OptDetallado Then
    arrform(5) = "tipo=' DETALLADO '"
    Call ImpresionRptProc("te_CajaBancoxRendir.rpt", arrform, arrparm)
   Else
          arrform(5) = "tipo=' RESUMIDO '"
          Call ImpresionRptProc("te_CajaBancoResumen.rpt", arrform, arrparm)
End If
End Sub

Sub ConfiguraCajaBanco(Valor As Boolean)
  Ctr_Ayu_Caja.Enabled = Valor
  Ctr_Ayu_Moneda.Enabled = Valor
  Ctr_Ayu_Caja.Visible = Valor
  Ctr_Ayu_Moneda.Visible = Valor
  
  lcaja.Visible = Valor
  lmon.Visible = Valor
  
  lban.Visible = Not Valor
  lcta.Visible = Not Valor
  Ctr_AyudaBanco.Enabled = Not Valor
  Ctr_AyudaBanco.Visible = Not Valor
  Ctr_AyudaBancoCuenta.Enabled = Not Valor
  Ctr_AyudaBancoCuenta.Visible = Not Valor
  
  If Valor = True Then
     Ctr_Ayu_Caja.ListaCampos = "cajacodigo(1),cajadescripcion(1)"
     Ctr_Ayu_Caja.ListaCamposDescrip = "Código,Descripción"
     Ctr_Ayu_Caja.ListaCamposText = "cajacodigo,cajadescripcion"
     Ctr_Ayu_Caja.NomTabla = "te_codigocaja"
     Ctr_Ayu_Caja.XcodCampo = "cajacodigo"
     Ctr_Ayu_Caja.XListCampo = "cajadescripcion"
     Ctr_Ayu_Caja.Conexion VGcnx
  Else
     Ctr_AyudaBanco.ListaCampos = "bancocodigo(1),bancodescripcion(1)"
     Ctr_AyudaBanco.ListaCamposDescrip = "Código,Descripción"
     Ctr_AyudaBanco.ListaCamposText = "bancocodigo,bancodescripcion"
     Ctr_AyudaBanco.NomTabla = "gr_banco"
     Ctr_AyudaBanco.XcodCampo = "bancocodigo"
     Ctr_AyudaBanco.XListCampo = "bancodescripcion"
     Ctr_AyudaBanco.Conexion VGcnx
  End If
  
  If Valor = True Then
      Ctr_Ayu_Moneda.ListaCampos = "monedacodigo(1),monedadescripcion(1)"
      Ctr_Ayu_Moneda.ListaCamposDescrip = "Código,Descripción"
      Ctr_Ayu_Moneda.ListaCamposText = "monedacodigo,monedadescripcion"
      Ctr_Ayu_Moneda.NomTabla = "gr_moneda"
      Ctr_Ayu_Moneda.XcodCampo = "monedacodigo"
      Ctr_Ayu_Moneda.XListCampo = "monedadescripcion"
      Ctr_Ayu_Moneda.Conexion VGcnx
  Else
      Ctr_AyudaBancoCuenta.ListaCampos = "cbanco_codigo(1),cbanco_numero(1),monedasimbolo(1),cbanco_referenciacta(1),cbanco_nrocheque(1),monedacodigo(1)"
      Ctr_AyudaBancoCuenta.ListaCamposDescrip = "Código,Descripción,Mon,Ref,NCheque,MonCod"
      Ctr_AyudaBancoCuenta.ListaCamposText = "cbanco_codigo,cbanco_numero,monedasimbolo,cbanco_referenciacta,cbanco_nrocheque,monedacodigo"
      Ctr_AyudaBancoCuenta.NomTabla = "v_bancomoneda"
      Ctr_AyudaBancoCuenta.XcodCampo = "cbanco_codigo"
      Ctr_AyudaBancoCuenta.XListCampo = "cbanco_numero"
      Ctr_AyudaBancoCuenta.Conexion VGcnx
  End If
  
End Sub

Sub ConfiguraBanco(Valor As Boolean)
  Ctr_AyudaBanco.Enabled = Valor
  Ctr_AyudaBancoCuenta.Enabled = Valor
End Sub

Private Sub Opt_Click(Index As Integer)
  Select Case Index
    Case 0:
       Call ConfiguraCajaBanco(True)
    
    Case 1:
       Call ConfiguraCajaBanco(False)
  End Select

End Sub


Private Sub OptCajaBancos_Click()
  FrameRendicion.Visible = False
  fraDetallado.Visible = False
If OptCajaBancos.Value = True Then
   fraDetallado.Visible = True
   FrameCajaBancos.Enabled = True
End If
End Sub

Private Sub OptRendiciones_Click()
  FrameRendicion.Visible = False
  fraDetallado.Visible = False

If OptRendiciones.Value = True Then
   FrameRendicion.Visible = True
   FrameCajaBancos.Enabled = False
End If
End Sub

