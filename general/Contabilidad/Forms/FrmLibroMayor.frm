VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Begin VB.Form FrmLibroMayor 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Libro Mayor"
   ClientHeight    =   4635
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   4635
   ScaleWidth      =   6585
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraDetallado 
      BackColor       =   &H00FFFFC0&
      Height          =   2385
      Left            =   240
      TabIndex        =   4
      Top             =   930
      Width           =   5955
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyudaSubAsiento 
         Height          =   315
         Left            =   1035
         TabIndex        =   5
         Top             =   1905
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
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyudaAsiento 
         Height          =   300
         Left            =   1035
         TabIndex        =   6
         Top             =   1575
         Width           =   4845
         _ExtentX        =   8546
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
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayucuenta 
         Height          =   360
         Index           =   0
         Left            =   1035
         TabIndex        =   7
         Top             =   780
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
         Left            =   4050
         TabIndex        =   8
         Top             =   270
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   529
         _Version        =   393216
         Format          =   87621633
         CurrentDate     =   37474
      End
      Begin MSComCtl2.DTPicker DTPickerFecInicio 
         Height          =   300
         Left            =   1035
         TabIndex        =   9
         Top             =   270
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   529
         _Version        =   393216
         Format          =   87621633
         CurrentDate     =   37474
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayucuenta 
         Height          =   360
         Index           =   1
         Left            =   1035
         TabIndex        =   10
         Top             =   1125
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
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_AyudaEntidad 
         Height          =   300
         Left            =   1035
         TabIndex        =   11
         Top             =   1575
         Visible         =   0   'False
         Width           =   4845
         _ExtentX        =   8546
         _ExtentY        =   529
         XcodMaxLongitud =   0
         xcodwith        =   1000
         NomTabla        =   "ct_entidad"
         TituloAyuda     =   "Ayuda Analítico"
         ListaCampos     =   "entidadcodigo(1),entidadrazonsocial(1)"
         XcodCampo       =   "entidadcodigo"
         XListCampo      =   "entidadrazonsocial"
         ListaCamposDescrip=   "Código,Razon_Social"
         ListaCamposText =   "entidadcodigo,entidadrazonsocial"
         Requerido       =   0   'False
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Fecha Final"
         Height          =   285
         Left            =   3150
         TabIndex        =   18
         Top             =   315
         Width           =   840
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Fecha Inicial"
         Height          =   300
         Left            =   75
         TabIndex        =   17
         Top             =   300
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Cuenta :"
         Height          =   255
         Left            =   75
         TabIndex        =   16
         Top             =   630
         Width           =   630
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Asiento"
         Height          =   255
         Left            =   105
         TabIndex        =   15
         Top             =   1635
         Width           =   930
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Sub Asiento"
         Height          =   285
         Left            =   75
         TabIndex        =   14
         Top             =   1980
         Width           =   885
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Desde"
         Height          =   195
         Left            =   495
         TabIndex        =   13
         Top             =   855
         Width           =   510
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Hasta"
         Height          =   240
         Left            =   495
         TabIndex        =   12
         Top             =   1215
         Width           =   465
      End
   End
   Begin VB.CommandButton cmdBotones 
      Caption         =   "&Cancelar"
      Height          =   375
      Index           =   1
      Left            =   3240
      TabIndex        =   3
      Top             =   3690
      Width           =   1320
   End
   Begin VB.CommandButton cmdBotones 
      Caption         =   "&Aceptar"
      Height          =   375
      Index           =   0
      Left            =   1560
      TabIndex        =   2
      Top             =   3705
      Width           =   1320
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   5955
      Begin VB.CheckBox chkAcumula 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Acumulado"
         Height          =   270
         Left            =   210
         TabIndex        =   1
         Top             =   180
         Width           =   1725
      End
   End
End
Attribute VB_Name = "FrmLibroMayor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public m_tituloreporte As String
Public m_caso As String

Private Sub chkAcumula_Click()
  If chkAcumula.Value = 0 Then
     Call SeleccionarMes(CInt(VGParamSistem.Mesproceso), CInt(VGParamSistem.Anoproceso))
     DTPickerFecInicio.Enabled = False
     DTPickerFecFinal.Enabled = False
  Else
    DTPickerFecInicio.Enabled = True
    DTPickerFecFinal.Enabled = True
  End If
End Sub

Private Sub Ctr_AyudaAsiento_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
  Ctr_AyudaSubAsiento.Filtro = "asientocodigo='" & Ctr_AyudaAsiento.xclave & "'"
End Sub

Private Sub Form_Load()
  Call ConfiguraForm
  Call SeleccionarMes(CInt(VGParamSistem.Mesproceso), CInt(VGParamSistem.Anoproceso))
  DTPickerFecInicio.Enabled = False
  DTPickerFecFinal.Enabled = False
End Sub

Sub ConfiguraForm()
  Me.Caption = m_tituloreporte
  Ctr_Ayucuenta(0).conexion VGCNx
  Ctr_Ayucuenta(1).conexion VGCNx
  Ctr_AyudaAsiento.conexion VGCNx
  Ctr_AyudaSubAsiento.conexion VGCNx
  Ctr_AyudaEntidad.conexion VGCNx
  Ctr_Ayucuenta(0).Filtro = "empresacodigo='" & VGParametros.empresacodigo & "' and cuentacodigo<>'00'"
  Ctr_Ayucuenta(1).Filtro = "empresacodigo='" & VGParametros.empresacodigo & "' and cuentacodigo<>'00'"
  Ctr_AyudaAsiento.Filtro = "asientocodigo<>'00'"
  Ctr_AyudaSubAsiento.Filtro = "subasientocodigo<>'00'"
  If m_caso = "1" Then
     Ctr_AyudaAsiento.Visible = False
     Ctr_AyudaSubAsiento.Visible = False
     Label4.Visible = False: Label5.Caption = " Analitico"
     Ctr_AyudaEntidad.Visible = True
  Else
     Ctr_AyudaAsiento.Visible = True
     Ctr_AyudaSubAsiento.Visible = True
     Label4.Visible = True
     Label5.Visible = True
     
     Ctr_AyudaEntidad.Visible = False
  End If
End Sub

Property Let tituloreporte(valor As String)
  m_tituloreporte = valor
End Property

Property Let Caso(valor As String)
  m_caso = valor
End Property

Private Sub cmdBotones_Click(Index As Integer)
 Select Case Index
  Case 0:
     Select Case m_caso
        Case "1":  'Imprimir Libro Mayor Analítico
          If ValidarImpresion = True Then
             Call ImpresionMayorAnalitico
          End If
        Case "2":  'Imprimir Libro Mayor General
            Call ImpresionMayorGeneral
     End Select
  
  Case 1: Unload Me
 
 End Select

End Sub

Sub SeleccionarMes(nMes As Integer, nAnno As Integer)
  DTPickerFecInicio.Value = Format("01/" & nMes & "/" & nAnno, "dd/mm/yyyy")
  DTPickerFecFinal.Value = DateAdd("d", -1, DateAdd("m", 1, DTPickerFecInicio.Value))
End Sub

Function ValidarImpresion() As Boolean
  If Format(DTPickerFecInicio.Value, "dd/mm/yyyy") > Format(DTPickerFecFinal.Value, "dd/mm/yyyy") Then
    MsgBox "La fecha de Término no puede ser mayor a la Fecha de Inicio", vbInformation, Caption
    ValidarImpresion = False
    DTPickerFecInicio.SetFocus
    Exit Function
  End If

  ValidarImpresion = True
End Function

Sub ImpresionMayorAnalitico()
'FIXIT: Declare 'arrform' and 'arrparm' con un tipo de datos de enlace en tiempo de compilación     FixIT90210ae-R1672-R1B8ZE
Dim arrform(2) As Variant, arrparm() As Variant
    ReDim arrparm(11)
    arrparm(0) = VGParamSistem.BDEmpresa
    arrparm(1) = VGParametros.empresacodigo
    arrparm(2) = VGParamSistem.Anoproceso
    arrparm(3) = Format(Month(DTPickerFecInicio.Value) - 1, "0#")
    arrparm(4) = Format(Month(DTPickerFecInicio.Value), "0#")
    arrparm(5) = Format(Month(DTPickerFecFinal.Value), "0#")
    arrparm(6) = IIf(Ctr_Ayucuenta(0).xclave = Empty, "%%", Trim$(Ctr_Ayucuenta(0).xclave))
    If Ctr_Ayucuenta(0).xclave = Ctr_Ayucuenta(1).xclave Then
        arrparm(7) = "%%"
    Else
        arrparm(7) = IIf(Ctr_Ayucuenta(1).xclave = Empty, "%%", Trim$(Ctr_Ayucuenta(1).xclave))
    End If
    arrparm(8) = IIf(Ctr_AyudaEntidad.xclave = Empty, "%%", Trim$(Ctr_AyudaEntidad.xclave) & "%")
    If Ctr_AyudaEntidad.xclave = Empty Then
        arrparm(9) = "TODOS"
    Else
        arrparm(9) = Ctr_AyudaEntidad.xnombre
    End If
    arrparm(10) = "FORMATO 06.01"
    Set VGvardllgen = New dllgeneral.dll_general
    arrform(0) = "@TituloReporte='" & "Libro Mayor Analítico Cuenta: " & "'"   'Ctr_Ayuda1.xclave
    If Month(DTPickerFecInicio.Value) <> Month(DTPickerFecFinal.Value) And Year(DTPickerFecInicio.Value) = Year(DTPickerFecFinal.Value) Then
       arrform(1) = "@Mes='" & VGvardllgen.DesMes(Month(DTPickerFecInicio.Value)) & " - " & VGvardllgen.DesMes(Month(DTPickerFecFinal.Value)) & "'"
    Else
       arrform(1) = "@Mes='" & VGvardllgen.DesMes(Month(DTPickerFecInicio.Value)) & "'"
    End If
    Call ImpresionRptProc("ct_libromayor.rpt", arrform, arrparm, , "Libro Mayor ")
End Sub

Sub ImpresionMayorGeneral()
'FIXIT: Declare 'arrform' and 'arrparm' con un tipo de datos de enlace en tiempo de compilación     FixIT90210ae-R1672-R1B8ZE
    Dim arrform(2) As Variant, arrparm() As Variant
    ReDim arrparm(10)
    arrparm(0) = VGParamSistem.BDEmpresa
    arrparm(1) = VGParametros.empresacodigo
    arrparm(2) = VGParamSistem.Anoproceso
    If chkAcumula.Value = 1 Then
        arrparm(3) = Format(Month(DTPickerFecInicio.Value), "00")
        arrparm(4) = Format(Month(DTPickerFecFinal.Value), "00")
        arrparm(9) = 1
    Else
        arrparm(3) = Format(VGParamSistem.Mesproceso - 1, "00")
        arrparm(4) = Format(VGParamSistem.Mesproceso, "00")
        arrparm(9) = 0
    End If
    arrparm(5) = IIf(Ctr_Ayucuenta(0).xclave = Empty, "%%", Trim$(Ctr_Ayucuenta(0).xclave))
    arrparm(6) = IIf(Ctr_AyudaAsiento.xclave = Empty, "%%", Ctr_AyudaAsiento.xclave)
    arrparm(7) = IIf(Ctr_AyudaSubAsiento.xclave = Empty, "%%", Ctr_AyudaSubAsiento.xclave)
    arrparm(8) = IIf(Ctr_Ayucuenta(1).xclave = Empty, "%%", Trim$(Ctr_Ayucuenta(1).xclave))
    Set VGvardllgen = New dllgeneral.dll_general
    arrform(0) = "@TituloReporte='" & "Libro Mayor General " & Ctr_Ayucuenta(0).xclave & " " & Ctr_Ayucuenta(0).xnombre & "'"
    arrform(1) = "@Mes='" & VGvardllgen.DesMes(VGParamSistem.Mesproceso) & " - " & VGParamSistem.Anoproceso & "'"
    Call ImpresionRptProc("ct_MayorGeneral.rpt", arrform, arrparm)
End Sub

