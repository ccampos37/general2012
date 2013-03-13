VERSION 5.00
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Begin VB.Form frmRepCtaCteAnalitico 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cuenta Corriente Entidades (Analíticos)"
   ClientHeight    =   6165
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7275
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   7275
   Begin VB.CheckBox chkAjuste 
      Caption         =   "Incluye ajuste por diferencia de cambio"
      Height          =   435
      Left            =   5370
      TabIndex        =   26
      Top             =   3510
      Width           =   1815
   End
   Begin VB.Frame Frame6 
      Caption         =   "Seleccionar Detalle"
      Height          =   690
      Left            =   90
      TabIndex        =   23
      Top             =   3330
      Width           =   5190
      Begin VB.OptionButton OptDetalle 
         Caption         =   "Solo Saldos"
         Height          =   195
         Index           =   0
         Left            =   225
         TabIndex        =   25
         Top             =   315
         Width           =   1410
      End
      Begin VB.OptionButton OptDetalle 
         Caption         =   "Detalle por documento"
         Height          =   240
         Index           =   1
         Left            =   2610
         TabIndex        =   24
         Top             =   270
         Value           =   -1  'True
         Width           =   1995
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Ordenado por"
      Height          =   855
      Left            =   4920
      TabIndex        =   18
      Top             =   120
      Width           =   2205
      Begin VB.OptionButton optOpcion 
         Caption         =   "Cuenta Contable"
         Height          =   225
         Index           =   3
         Left            =   135
         TabIndex        =   20
         Top             =   540
         Width           =   3120
      End
      Begin VB.OptionButton optOpcion 
         Caption         =   "Entidad"
         Height          =   225
         Index           =   2
         Left            =   135
         TabIndex        =   19
         Top             =   285
         Width           =   3000
      End
   End
   Begin VB.ListBox lstTipoAnaliticoCodigo 
      Height          =   255
      Left            =   5805
      TabIndex        =   14
      Top             =   4950
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.CommandButton cmdBotones 
      Caption         =   "&Aceptar"
      Height          =   375
      Index           =   0
      Left            =   2085
      TabIndex        =   13
      Top             =   4980
      Width           =   1470
   End
   Begin VB.CommandButton cmdBotones 
      Caption         =   "&Cancelar"
      Height          =   375
      Index           =   1
      Left            =   3765
      TabIndex        =   12
      Top             =   4980
      Width           =   1470
   End
   Begin VB.Frame Frame4 
      Caption         =   "Seleccionar Filtro"
      Height          =   750
      Left            =   90
      TabIndex        =   11
      Top             =   4095
      Width           =   7125
      Begin VB.OptionButton optFiltro 
         Caption         =   "Todos"
         Height          =   300
         Index           =   2
         Left            =   5445
         TabIndex        =   17
         Top             =   270
         Width           =   1515
      End
      Begin VB.OptionButton optFiltro 
         Caption         =   "Documentos pendientes"
         Height          =   300
         Index           =   1
         Left            =   2595
         TabIndex        =   16
         Top             =   270
         Width           =   2145
      End
      Begin VB.OptionButton optFiltro 
         Caption         =   "Documentos cancelados"
         Height          =   300
         Index           =   0
         Left            =   180
         TabIndex        =   15
         Top             =   270
         Width           =   2145
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Seleccionar Cuentas"
      Height          =   945
      Left            =   45
      TabIndex        =   9
      Top             =   1890
      Width           =   7140
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuda1 
         Height          =   360
         Left            =   1680
         TabIndex        =   5
         Top             =   360
         Width           =   5385
         _ExtentX        =   9499
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
      Begin VB.Label Label1 
         Caption         =   "Cuenta Contable"
         Height          =   285
         Left            =   135
         TabIndex        =   10
         Top             =   435
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Seleccionar Entidades (Analítico)"
      Height          =   720
      Left            =   30
      TabIndex        =   3
      Top             =   1125
      Width           =   7110
      Begin VB.ComboBox cboTipoAnalitico 
         Height          =   315
         Left            =   1890
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   285
         Width           =   4395
      End
      Begin VB.Label Label2 
         Caption         =   "Tipo Analítico"
         Height          =   240
         Left            =   180
         TabIndex        =   8
         Top             =   345
         Width           =   1005
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Emitir por"
      Height          =   855
      Left            =   45
      TabIndex        =   2
      Top             =   105
      Width           =   4725
      Begin VB.OptionButton optOpcion 
         Caption         =   "Saldos iniciales"
         Height          =   225
         Index           =   4
         Left            =   2640
         TabIndex        =   21
         Top             =   360
         Width           =   3000
      End
      Begin VB.OptionButton optOpcion 
         Caption         =   "Movimiento Cuenta Corriente"
         Height          =   225
         Index           =   0
         Left            =   135
         TabIndex        =   0
         Top             =   285
         Width           =   3000
      End
      Begin VB.OptionButton optOpcion 
         Caption         =   "Saldos Cuenta Corriente"
         Height          =   225
         Index           =   1
         Left            =   135
         TabIndex        =   1
         Top             =   555
         Width           =   3000
      End
   End
   Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyudaEnt 
      Height          =   330
      Left            =   1665
      TabIndex        =   7
      Top             =   2925
      Width           =   5385
      _ExtentX        =   9499
      _ExtentY        =   582
      XcodMaxLongitud =   0
      xcodwith        =   1300
      NomTabla        =   "ct_entidad"
      TituloAyuda     =   "Ayuda Búsqueda Entidad"
      ListaCampos     =   "entidadcodigo(1),entidadrazonsocial(1)"
      XcodCampo       =   "entidadcodigo"
      XListCampo      =   "entidadrazonsocial"
      ListaCamposDescrip=   "Código,Razon_Social"
      ListaCamposText =   "entidadcodigo,entidadrazonsocial"
      Requerido       =   0   'False
   End
   Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuda2 
      Height          =   330
      Left            =   1665
      TabIndex        =   6
      Top             =   2925
      Visible         =   0   'False
      Width           =   5385
      _ExtentX        =   9499
      _ExtentY        =   582
      XcodMaxLongitud =   0
      xcodwith        =   1300
      NomTabla        =   "v_analiticoentidad"
      ListaCampos     =   "analiticocodigo(1),entidadrazonsocial(1)"
      XcodCampo       =   "analiticocodigo"
      XListCampo      =   "entidadrazonsocial"
      ListaCamposDescrip=   "Código,Razon_Social"
      ListaCamposText =   "analiticocodigo,entidadrazonsocial"
      Requerido       =   0   'False
   End
   Begin VB.Label Label4 
      Caption         =   "Código Analítico"
      Height          =   270
      Left            =   180
      TabIndex        =   22
      Top             =   2970
      Width           =   1500
   End
End
Attribute VB_Name = "frmRepCtaCteAnalitico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
   Call LlenarcboTipoAnalitico
   Ctr_Ayuda1.conexion VGCNx
   Ctr_Ayuda2.conexion VGCNx
   Ctr_AyudaEnt.conexion VGCNx
   Ctr_Ayuda1.Filtro = " empresacodigo='" & VGParametros.empresacodigo & "'"
   Me.Height = 5970
   Me.Width = 7320
   optFiltro(0).Value = True
End Sub

Private Sub cboTipoAnalitico_Click()
  Ctr_Ayuda2.Filtro = "Right(tipoanaliticocodigo,3)='" & lstTipoAnaliticoCodigo.List(cboTipoAnalitico.ListIndex) & "'"
  Ctr_Ayuda2.Ejecutar
  
  Ctr_AyudaEnt.Filtro = "entidadcodigo in (Select case when lTrim(rtrim(entidadcodigo))='' then '0' else entidadcodigo end as entidadcodigo From ct_analitico Where tipoanaliticocodigo = '" & lstTipoAnaliticoCodigo.List(cboTipoAnalitico.ListIndex) & "')"
  Ctr_AyudaEnt.Ejecutar
End Sub

Private Sub cmdBotones_Click(Index As Integer)
   Select Case Index
     Case 0
       If ValidaImpresion = True Then
            If optOpcion(0).Value = True Then
                Call ImpresionCtacteMovimiento
             ElseIf optOpcion(1).Value = True Then
                   Call ImpresionCtaCteSaldo
                 ElseIf optOpcion(4).Value = True Then
                   Call ImpresionCtaCteSaldoinicial
                   
            End If
       End If
     
     Case 1: Unload Me
   End Select
End Sub

Function ValidaImpresion() As Boolean
  If cboTipoAnalitico.ListIndex < 0 Then
     If Not MsgBox("No ha seleccionado el tipo de anallitico, desea continuar ", vbYesNo + vbDefaultButton2, "AVISO") = vbYes Then
        ValidaImpresion = False
        Exit Function
     End If
  End If
  ValidaImpresion = True
End Function

Sub LlenarcboTipoAnalitico()
   Dim dllgen As New dllgeneral.dll_general
   Dim rs As ADODB.Recordset
   
   Set rs = VGCNx.Execute("Select tipoanaliticocodigo,tipoanaliticodescripcion from ct_tipoanalitico where tipoanaliticocodigo<>'00'")
   cboTipoAnalitico.Clear
   lstTipoAnaliticoCodigo.Clear
   While Not rs.EOF
     cboTipoAnalitico.AddItem rs(1)
     lstTipoAnaliticoCodigo.AddItem rs(0)
     rs.MoveNext
   Wend
   Set dllgen = New dllgeneral.dll_general
End Sub

Sub ImpresionCtacteMovimiento()
Dim cMensaje As String
'FIXIT: Declare 'arrform' and 'arrparm' con un tipo de datos de enlace en tiempo de compilación     FixIT90210ae-R1672-R1B8ZE
Dim arrform(2) As Variant, arrparm() As Variant
    ReDim arrparm(14)
      
    arrparm(0) = VGParamSistem.BDEmpresa
    arrparm(1) = VGParametros.empresacodigo
    arrparm(2) = VGParamSistem.Anoproceso
    arrparm(3) = "00"
    arrparm(4) = Format(VGParamSistem.Mesproceso, "0#")
    arrparm(5) = IIf(Ctr_Ayuda1.xclave = Empty, "%%", Trim(Ctr_Ayuda1.xclave) & "%")
    arrparm(6) = "%%"
    arrparm(7) = "%%"
    
    arrparm(8) = IIf(Ctr_AyudaEnt.xclave = Empty, "%%", Trim(Ctr_AyudaEnt.xclave) & "%")
    
    arrparm(9) = IIf(lstTipoAnaliticoCodigo.List(cboTipoAnalitico.ListIndex) = Empty, "%%", lstTipoAnaliticoCodigo.List(cboTipoAnalitico.ListIndex))
    If optFiltro(0).Value = True Then
        arrparm(10) = "1"
        cMensaje = "Cancelados"
    ElseIf optFiltro(1).Value = True Then
        arrparm(10) = "2"
        cMensaje = "Pendientes"
    Else
        arrparm(10) = "3"
        cMensaje = "Todos"
    End If

    If OptDetalle(0).Value = True Then
        If Ctr_AyudaEnt.Visible = True And Ctr_AyudaEnt.xclave <> Empty And Ctr_Ayuda1.xclave = Empty Then
            optOpcion(2).Value = True
            arrparm(11) = 1
        ElseIf Ctr_Ayuda1.xclave <> Empty Then
            optOpcion(3).Value = True
            arrparm(11) = 2
        Else
            arrparm(11) = 0
        End If
    Else
        arrparm(11) = 0
    End If
    arrparm(12) = 0
    arrparm(13) = IIf(chkAjuste.Value = 1, "%", "0")  ' param=0 -> No incluye asientos de ajuste x dif. cambio
    Set VGvardllgen = New dllgeneral.dll_general
    arrform(0) = "@TituloReporte='" & "Reporte de Movimientos de Cuenta Corriente de " & cboTipoAnalitico.Text & " '"
    arrform(1) = "@Mes='" & VGvardllgen.DesMes(VGParamSistem.Mesproceso) & " - " & VGParamSistem.Anoproceso & " - Situación: " & cMensaje & "'"
    arrform(2) = "@MonedaBase='" & VGParametros.monedabase & "'"
    If optOpcion(2).Value = True Then
       Call ImpresionRptProc("ct_CtacteAnalitico1.rpt", arrform, arrparm)
     Else
        If OptDetalle(0).Value = True Then
            Call ImpresionRptProc("ct_CtacteAnaliticoSaldos.rpt", arrform, arrparm)
        Else
            Call ImpresionRptProc("ct_CtacteAnalitico2.rpt", arrform, arrparm)
        End If
    End If
End Sub

Sub ImpresionCtaCteSaldo()
'FIXIT: Declare 'arrform' and 'arrparm' con un tipo de datos de enlace en tiempo de compilación     FixIT90210ae-R1672-R1B8ZE
Dim arrform(2) As Variant, arrparm() As Variant
    ReDim arrparm(9)
    arrparm(0) = VGParamSistem.BDEmpresa
    arrparm(1) = VGParametros.empresacodigo
    arrparm(2) = VGParamSistem.Anoproceso
    arrparm(3) = "00"
    arrparm(4) = Format(VGParamSistem.Mesproceso, "0#")
    arrparm(5) = IIf(Ctr_Ayuda1.xclave = Empty, "%%", Trim(Ctr_Ayuda1.xclave) & "%")
    arrparm(6) = IIf(Ctr_AyudaEnt.xclave = Empty, "%%", Trim(Ctr_AyudaEnt.xclave) & "%")
    arrparm(7) = IIf(lstTipoAnaliticoCodigo.List(cboTipoAnalitico.ListIndex) = Empty, "%%", lstTipoAnaliticoCodigo.List(cboTipoAnalitico.ListIndex))
    arrparm(8) = IIf(chkAjuste.Value = 1, "1", "0")  ' param=0 -> No toma en cuenta dif. cambio
    
    Set VGvardllgen = New dllgeneral.dll_general
    arrform(0) = "@TituloReporte='" & "Saldos de Cuenta Corriente " & cboTipoAnalitico.Text & " al Mes de " & "'"
    arrform(1) = "@Mes='" & VGvardllgen.DesMes(VGParamSistem.Mesproceso) & "'"
    If optOpcion(2).Value = True Then
       Call ImpresionRptProc("rptSaldoAnalitico1.rpt", arrform, arrparm)
     Else
       Call ImpresionRptProc("rptSaldoAnalitico2.rpt", arrform, arrparm)
    End If
End Sub
Sub ImpresionCtaCteSaldoinicial()
Dim cMensaje As String
'FIXIT: Declare 'arrform' and 'arrparm' con un tipo de datos de enlace en tiempo de compilación     FixIT90210ae-R1672-R1B8ZE
Dim arrform(2) As Variant, arrparm() As Variant
If optFiltro(1).Value = True Then
    ReDim arrparm(9)
    arrparm(0) = VGParamSistem.BDEmpresa
    arrparm(1) = VGParametros.empresacodigo
    arrparm(2) = VGParamSistem.Anoproceso
    arrparm(3) = "00"
    arrparm(4) = "00"
    arrparm(5) = IIf(Ctr_Ayuda1.xclave = Empty, "%%", Trim(Ctr_Ayuda1.xclave) & "%")
    arrparm(6) = IIf(Ctr_AyudaEnt.xclave = Empty, "%%", Trim(Ctr_AyudaEnt.xclave) & "%")
    arrparm(7) = IIf(lstTipoAnaliticoCodigo.List(cboTipoAnalitico.ListIndex) = Empty, "%%", lstTipoAnaliticoCodigo.List(cboTipoAnalitico.ListIndex))
    arrparm(8) = IIf(chkAjuste.Value = 1, "1", "0")  ' param=0 -> No toma en cuenta dif. cambio
        Call ImpresionRptProc("rptSaldoAnalitico2.rpt", arrform, arrparm)
Else
    ReDim arrparm(14)
    arrparm(0) = VGParamSistem.BDEmpresa
    arrparm(1) = VGParametros.empresacodigo
    arrparm(2) = VGParamSistem.Anoproceso
    arrparm(3) = "00"
    arrparm(4) = "00"
    arrparm(5) = IIf(Ctr_Ayuda1.xclave = Empty, "%%", Trim(Ctr_Ayuda1.xclave) & "%")
    arrparm(6) = "%%"
    arrparm(7) = "%%"
    arrparm(8) = IIf(Ctr_AyudaEnt.xclave = Empty, "%%", Trim(Ctr_AyudaEnt.xclave) & "%")
    arrparm(9) = IIf(lstTipoAnaliticoCodigo.List(cboTipoAnalitico.ListIndex) = Empty, "%%", lstTipoAnaliticoCodigo.List(cboTipoAnalitico.ListIndex))
'    arrparm(10) = "0"
    arrparm(10) = "3"
    cMensaje = "Saldos Iniciales"
    Set VGvardllgen = New dllgeneral.dll_general
    arrform(0) = "@TituloReporte='" & "Reporte de Saldos Iniciales de Cuenta Corriente de " & cboTipoAnalitico.Text & " '"
    arrform(1) = "@Mes='Ano  - " & VGParamSistem.Anoproceso & " - Situación: " & cMensaje & "'"
    arrform(2) = "@MonedaBase='" & VGParametros.monedabase & "'"
    arrparm(11) = 0
    arrparm(12) = 0
    arrparm(13) = 0  ' param=0 -> No incluye asientos de ajuste x dif. cambio
    
    Call ImpresionRptProc("ct_CtacteAnalitico2.rpt", arrform, arrparm)
End If

End Sub

Private Sub optOpcion_Click(Index As Integer)
    If Index = 0 Then
        OptDetalle(0).Enabled = True
        OptDetalle(1).Enabled = True
    Else
        OptDetalle(0).Enabled = False
        OptDetalle(1).Enabled = False
    End If
End Sub
